use serde::{Deserialize, Serialize};
use std::process::Command;
use std::os::windows::process::CommandExt;
use std::sync::Mutex;
use tauri::State;
use std::path::{Path, PathBuf};
use std::fs;
use std::time::{SystemTime, UNIX_EPOCH};

const CREATE_NO_WINDOW: u32 = 0x08000000;

#[derive(Default)]
struct AppState {
    current_engine: Mutex<String>,
}

#[derive(Serialize)]
struct EngineStatus {
    available: bool,
    engine: String,
    all_engines: Vec<String>,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
struct BatchResult {
    path: String,
    success: bool,
    pdf_path: Option<String>,
    output_path: Option<String>,
    error_msg: Option<String>,
}

#[tauri::command]
fn check_engines(state: State<'_, AppState>) -> Result<EngineStatus, String> {
    let mut all_engines = Vec::new();

    let office_ppt = Command::new("reg").args(["query", r#"HKCR\PowerPoint.Application"#, "/ve"]).creation_flags(CREATE_NO_WINDOW).output().map(|o| o.status.success()).unwrap_or(false);
    let office_word = Command::new("reg").args(["query", r#"HKCR\Word.Application"#, "/ve"]).creation_flags(CREATE_NO_WINDOW).output().map(|o| o.status.success()).unwrap_or(false);
    let office_excel = Command::new("reg").args(["query", r#"HKCR\Excel.Application"#, "/ve"]).creation_flags(CREATE_NO_WINDOW).output().map(|o| o.status.success()).unwrap_or(false);

    if office_ppt || office_word || office_excel {
        all_engines.push("office".to_string());
    }

    let wps_ppt = Command::new("reg").args(["query", r#"HKCR\KWPP.Application"#, "/ve"]).creation_flags(CREATE_NO_WINDOW).output().map(|o| o.status.success()).unwrap_or(false);
    let wps_word = Command::new("reg").args(["query", r#"HKCR\KWPS.Application"#, "/ve"]).creation_flags(CREATE_NO_WINDOW).output().map(|o| o.status.success()).unwrap_or(false);
    let wps_excel = Command::new("reg").args(["query", r#"HKCR\KET.Application"#, "/ve"]).creation_flags(CREATE_NO_WINDOW).output().map(|o| o.status.success()).unwrap_or(false);

    if wps_ppt || wps_word || wps_excel {
        all_engines.push("wps".to_string());
    }

    let lo_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ];
    for p in lo_paths.iter() {
        if Path::new(p).exists() {
            all_engines.push("libreoffice".to_string());
            break;
        }
    }

    if all_engines.is_empty() {
        return Ok(EngineStatus { available: false, engine: "".to_string(), all_engines: vec![] });
    }

    let selected = all_engines[0].clone();
    *state.current_engine.lock().unwrap() = selected.clone();

    Ok(EngineStatus { available: true, engine: selected, all_engines })
}

#[tauri::command]
fn set_engine(engine: String, state: State<'_, AppState>) -> Result<bool, String> {
    *state.current_engine.lock().unwrap() = engine;
    Ok(true)
}

#[tauri::command]
async fn convert_batch(paths: Vec<String>, state: State<'_, AppState>) -> Result<Vec<BatchResult>, String> {
    let engine = state.current_engine.lock().unwrap().clone();
    let current_dir = std::env::temp_dir();
    
    let mut word_files = Vec::new();
    let mut excel_files = Vec::new();
    let mut ppt_files = Vec::new();
    let mut libre_files = Vec::new();

    for path in paths {
        let p = PathBuf::from(&path);
        let ext = p.extension().and_then(|e| e.to_str()).unwrap_or("").to_lowercase();
        
        if engine == "libreoffice" {
            libre_files.push(path.clone());
            continue;
        }
        if ext == "doc" || ext == "docx" {
            word_files.push(path);
        } else if ext == "xls" || ext == "xlsx" || ext == "csv" {
            excel_files.push(path);
        } else {
            ppt_files.push(path);
        }
    }

    let mut results = Vec::new();

    let run_ps = |script: String, res_path: &Path| -> Vec<BatchResult> {
        let timestamp = SystemTime::now().duration_since(UNIX_EPOCH).unwrap().as_millis();
        let script_path = current_dir.join(format!("batch_{}.ps1", timestamp));
        let mut content = vec![0xEF, 0xBB, 0xBF];
        content.extend_from_slice(script.as_bytes());
        if fs::write(&script_path, content).is_err() {
            return vec![]; 
        }

        let _ = Command::new("powershell")
            .creation_flags(CREATE_NO_WINDOW)
            .args(["-ExecutionPolicy", "Bypass", "-NoProfile", "-WindowStyle", "Hidden", "-File", script_path.to_str().unwrap()])
            .output();
            
        let _ = fs::remove_file(&script_path);

        let mut res = Vec::new();
        if let Ok(json_str) = fs::read_to_string(res_path) {
            let clean_str = json_str.trim_start_matches('\u{feff}');
            // PowerShell Pipeline unwraps single-length arrays into objects, so handle both Array and Single object cases
            if let Ok(parsed) = serde_json::from_str::<Vec<BatchResult>>(clean_str) {
                res = parsed;
            } else if let Ok(parsed_single) = serde_json::from_str::<BatchResult>(clean_str) {
                res.push(parsed_single);
            }
            let _ = fs::remove_file(res_path);
        }
        res
    };

    if !word_files.is_empty() {
        let app_id = if engine == "office" { "Word.Application" } else { "KWPS.Application" };
        let mut ps_array = String::new();
        for f in &word_files { ps_array.push_str(&format!("'{}',", f.replace("'", "''"))); }
        if !ps_array.is_empty() { ps_array.pop(); }
        
        let res_path = current_dir.join(format!("res_word_{}.json", SystemTime::now().duration_since(UNIX_EPOCH).unwrap().as_millis()));
        let res_path_str = res_path.to_string_lossy().replace("'", "''");

        let script = format!(r#"
$ErrorActionPreference = "Continue"
$results = @()
try {{ $app = New-Object -ComObject {} }} catch {{ exit 1 }}
$app.Visible = $false
$files = @({})
foreach ($file in $files) {{
    try {{
        $doc = $app.Documents.Open($file, $false, $true)
        $pdfPath = [System.IO.Path]::ChangeExtension($file, '.pdf')
        $doc.ExportAsFixedFormat($pdfPath, 17)
        $doc.Close($false)
        $results += [PSCustomObject] @{{ path = $file; success = $true; pdf_path = $pdfPath; error_msg = $null }}
    }} catch {{
        $results += [PSCustomObject] @{{ path = $file; success = $false; pdf_path = $null; error_msg = $($_.Exception.Message) }}
    }}
}}
try {{ $app.Quit() }} catch {{}}
try {{ [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null }} catch {{}}
[System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
$results | ConvertTo-Json -Depth 3 | Out-File -FilePath '{}' -Encoding UTF8
"#, app_id, ps_array, res_path_str);
        results.extend(run_ps(script, &res_path));
    }

    if !excel_files.is_empty() {
        let app_id = if engine == "office" { "Excel.Application" } else { "KET.Application" };
        let mut ps_array = String::new();
        for f in &excel_files { ps_array.push_str(&format!("'{}',", f.replace("'", "''"))); }
        if !ps_array.is_empty() { ps_array.pop(); }
        
        let res_path = current_dir.join(format!("res_excel_{}.json", SystemTime::now().duration_since(UNIX_EPOCH).unwrap().as_millis()));
        let res_path_str = res_path.to_string_lossy().replace("'", "''");

        let script = format!(r#"
$ErrorActionPreference = "Continue"
$results = @()
try {{ $app = New-Object -ComObject {} }} catch {{ exit 1 }}
$app.Visible = $false
$app.DisplayAlerts = $false
$files = @({})
foreach ($file in $files) {{
    try {{
        $doc = $app.Workbooks.Open($file, 0, $true)
        $pdfPath = [System.IO.Path]::ChangeExtension($file, '.pdf')
        $doc.ExportAsFixedFormat(0, $pdfPath, 0, $true, $false)
        $doc.Close($false)
        $results += [PSCustomObject] @{{ path = $file; success = $true; pdf_path = $pdfPath; error_msg = $null }}
    }} catch {{
        $results += [PSCustomObject] @{{ path = $file; success = $false; pdf_path = $null; error_msg = $($_.Exception.Message) }}
    }}
}}
try {{ $app.Quit() }} catch {{}}
try {{ [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null }} catch {{}}
[System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
$results | ConvertTo-Json -Depth 3 | Out-File -FilePath '{}' -Encoding UTF8
"#, app_id, ps_array, res_path_str);
        results.extend(run_ps(script, &res_path));
    }

    if !ppt_files.is_empty() {
        let app_id = if engine == "office" { "PowerPoint.Application" } else { "KWPP.Application" };
        let mut ps_array = String::new();
        for f in &ppt_files { ps_array.push_str(&format!("'{}',", f.replace("'", "''"))); }
        if !ps_array.is_empty() { ps_array.pop(); }
        
        let res_path = current_dir.join(format!("res_ppt_{}.json", SystemTime::now().duration_since(UNIX_EPOCH).unwrap().as_millis()));
        let res_path_str = res_path.to_string_lossy().replace("'", "''");

        let script = format!(r#"
$ErrorActionPreference = "Continue"
$results = @()
try {{ $app = New-Object -ComObject {} }} catch {{ exit 1 }}
$files = @({})
foreach ($file in $files) {{
    try {{
        $pres = $app.Presentations.Open($file, $true, $false, $false)
        $pdfPath = [System.IO.Path]::ChangeExtension($file, '.pdf')
        $pres.SaveAs($pdfPath, 32)
        $pres.Close()
        $results += [PSCustomObject] @{{ path = $file; success = $true; pdf_path = $pdfPath; error_msg = $null }}
    }} catch {{
        $results += [PSCustomObject] @{{ path = $file; success = $false; pdf_path = $null; error_msg = $($_.Exception.Message) }}
    }}
}}
try {{ $app.Quit() }} catch {{}}
try {{ [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null }} catch {{}}
[System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
$results | ConvertTo-Json -Depth 3 | Out-File -FilePath '{}' -Encoding UTF8
"#, app_id, ps_array, res_path_str);
        results.extend(run_ps(script, &res_path));
    }
    
    // libreoffice handling is unaltered for simplicity
    if !libre_files.is_empty() {
        let mut soffice = String::new();
        let lo_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ];
        for p in lo_paths.iter() {
            if Path::new(p).exists() {
                soffice = p.to_string();
                break;
            }
        }
        
        for file in libre_files {
            if soffice.is_empty() {
                results.push(BatchResult { path: file.clone(), success: false, pdf_path: None, error_msg: Some("LibreOffice 未安装".to_string()) });
                continue;
            }
            let parent = Path::new(&file).parent().unwrap_or(Path::new("")).to_str().unwrap();
            let out = Command::new(&soffice).creation_flags(CREATE_NO_WINDOW).args(["--headless", "--convert-to", "pdf", "--outdir", parent, &file]).output();
            let pdf_path = PathBuf::from(&file).with_extension("pdf").to_string_lossy().to_string();
            
            if let Ok(o) = out {
                if o.status.success() {
                    results.push(BatchResult { path: file, success: true, pdf_path: Some(pdf_path), error_msg: None });
                } else {
                    results.push(BatchResult { path: file, success: false, pdf_path: None, error_msg: Some(String::from_utf8_lossy(&o.stderr).to_string()) });
                }
            } else {
                results.push(BatchResult { path: file, success: false, pdf_path: None, error_msg: Some("执行调用异常".to_string()) });
            }
        }
    }

    Ok(results)
}

#[tauri::command]
fn open_file(path: String) -> Result<(), String> {
    Command::new("cmd").creation_flags(CREATE_NO_WINDOW).args(["/C", "start", "", &path]).spawn().map_err(|e| e.to_string())?;
    Ok(())
}

#[tauri::command]
fn open_folder(path: String) -> Result<(), String> {
    let p = PathBuf::from(&path);
    if let Some(parent) = p.parent() {
        Command::new("cmd").creation_flags(CREATE_NO_WINDOW).args(["/C", "start", "", parent.to_str().unwrap()]).spawn().map_err(|e| e.to_string())?;
    }
    Ok(())
}

#[tauri::command]
async fn select_files(mode: String) -> Result<Vec<String>, String> {
    let mut dialog = rfd::AsyncFileDialog::new();
    if mode == "pdf2image" {
        dialog = dialog.add_filter("PDF Document", &["pdf"]);
    } else {
        dialog = dialog.add_filter("Office Documents", &["ppt", "pptx", "pps", "ppsx", "doc", "docx", "xls", "xlsx", "csv"]);
    }
    
    if let Some(files) = dialog.pick_files().await {
        Ok(files.into_iter().map(|f| f.path().to_string_lossy().into_owned()).collect())
    } else {
        Ok(vec![])
    }
}

#[tauri::command]
async fn convert_pdf_to_images(
    paths: Vec<String>,
    format: String,
    dpi: f32,
) -> Result<Vec<BatchResult>, String> {
    pdf_converter::convert_pdf_batch(paths, format, dpi).await
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .manage(AppState::default())
        .plugin(tauri_plugin_opener::init())
        .invoke_handler(tauri::generate_handler![
            check_engines,
            set_engine,
            convert_batch,
            convert_pdf_to_images,
            open_file,
            open_folder,
            select_files
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}

#[cfg(target_os = "windows")]
mod pdf_converter {
    use super::BatchResult;
    use std::path::Path;
    use windows::core::{HSTRING, GUID};
    use windows::Storage::StorageFile;
    use windows::Data::Pdf::{PdfDocument, PdfPageRenderOptions};
    use windows::Storage::Streams::{InMemoryRandomAccessStream, DataReader};
    use windows::Graphics::Imaging::BitmapEncoder;

    pub async fn convert_pdf_batch(
        paths: Vec<String>,
        format: String,
        dpi: f32,
    ) -> Result<Vec<BatchResult>, String> {
        let mut results = Vec::new();
        
        let encoder_guid = match format.to_lowercase().as_str() {
            "jpg" | "jpeg" => BitmapEncoder::JpegEncoderId().map_err(|e| e.to_string())?,
            _ => BitmapEncoder::PngEncoderId().map_err(|e| e.to_string())?,
        };

        for path_str in paths {
            match convert_single_pdf(&path_str, &format, encoder_guid, dpi).await {
                Ok(out_dir) => {
                    results.push(BatchResult {
                        path: path_str,
                        success: true,
                        pdf_path: None,
                        output_path: Some(out_dir),
                        error_msg: None,
                    });
                }
                Err(err_msg) => {
                    results.push(BatchResult {
                        path: path_str,
                        success: false,
                        pdf_path: None,
                        output_path: None,
                        error_msg: Some(err_msg),
                    });
                }
            }
        }
        Ok(results)
    }

    async fn convert_single_pdf(
        pdf_path_str: &str,
        format_ext: &str,
        encoder_guid: GUID,
        dpi: f32,
    ) -> Result<String, String> {
        let pdf_path = Path::new(pdf_path_str);
        if !pdf_path.exists() {
            return Err("文件不存在".to_string());
        }
        
        let file_stem = pdf_path.file_stem().ok_or("无效文件名")?.to_string_lossy().into_owned();
        let parent_dir = pdf_path.parent().ok_or("找不到父目录")?;
        
        let out_dir = parent_dir.join(format!("{}_images", file_stem));
        std::fs::create_dir_all(&out_dir).map_err(|e| format!("创建输出目录失败: {}", e))?;

        let abs_path = std::fs::canonicalize(pdf_path).map_err(|e| e.to_string())?;
        let win_path = abs_path.to_string_lossy().replace("\\\\?\\", "");

        let file = StorageFile::GetFileFromPathAsync(&HSTRING::from(&win_path))
            .map_err(|e| format!("加载PDF失败: {}", e))?
            .await
            .map_err(|e| format!("加载PDF失败: {}", e))?;

        let doc = PdfDocument::LoadFromFileAsync(&file)
            .map_err(|e| format!("打开PDF失败: {}", e))?
            .await
            .map_err(|e| format!("打开PDF失败: {}", e))?;

        let page_count = doc.PageCount().map_err(|e| e.to_string())?;

        for i in 0..page_count {
            let page = doc.GetPage(i).map_err(|e| e.to_string())?;
            let stream = InMemoryRandomAccessStream::new().map_err(|e| e.to_string())?;

            let options = PdfPageRenderOptions::new().map_err(|e| e.to_string())?;
            options.SetBitmapEncoderId(encoder_guid).map_err(|e| e.to_string())?;
            
            let size = page.Size().map_err(|e| e.to_string())?;
            let scale = dpi / 96.0;
            options.SetDestinationWidth((size.Width * scale) as u32).map_err(|e| e.to_string())?;
            options.SetDestinationHeight((size.Height * scale) as u32).map_err(|e| e.to_string())?;

            page.RenderToStreamWithOptionsAsync(&stream, &options)
                .map_err(|e| format!("渲染第 {} 页失败: {}", i + 1, e))?
                .await
                .map_err(|e| format!("渲染第 {} 页失败: {}", i + 1, e))?;

            let size_bytes = stream.Size().map_err(|e| e.to_string())?;
            let input_stream = stream.GetInputStreamAt(0).map_err(|e| e.to_string())?;
            let reader = DataReader::CreateDataReader(&input_stream).map_err(|e| e.to_string())?;
            
            reader.LoadAsync(size_bytes as u32)
                .map_err(|e| e.to_string())?
                .await
                .map_err(|e| e.to_string())?;

            let mut buffer = vec![0u8; size_bytes as usize];
            reader.ReadBytes(&mut buffer).map_err(|e| e.to_string())?;

            let out_img_name = format!("{}_{:03}.{}", file_stem, i + 1, format_ext.to_lowercase());
            std::fs::write(out_dir.join(out_img_name), buffer).map_err(|e| format!("写入图片失败: {}", e))?;
        }

        Ok(out_dir.to_string_lossy().into_owned())
    }
}

#[cfg(not(target_os = "windows"))]
mod pdf_converter {
    use super::BatchResult;
    pub async fn convert_pdf_batch(
        paths: Vec<String>,
        _format: String,
        _dpi: f32,
    ) -> Result<Vec<BatchResult>, String> {
        let mut results = Vec::new();
        for p in paths {
            results.push(BatchResult {
                path: p,
                success: false,
                pdf_path: None,
                output_path: None,
                error_msg: Some("PDF 转换图片功能目前仅在 Windows 系统下生效。".to_string()),
            });
        }
        Ok(results)
    }
}
