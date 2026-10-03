//! Run with `cargo run -p osbrowser --example dom`.

use serde_json::json;

fn main() -> Result<(), String> {
    let browser = osbrowser::BrowserSession::start()?;
    let args: Vec<String> = std::env::args().collect();
    let smoke = args.iter().any(|arg| arg == "--headless-smoke");
    let mut headless = None;
    let mut profile = None;
    if smoke {
        #[cfg(target_os = "macos")]
        let chrome = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
        #[cfg(target_os = "windows")]
        let chrome = "chrome.exe";
        #[cfg(all(unix, not(target_os = "macos")))]
        let chrome = "google-chrome";
        let path = std::env::temp_dir().join(format!(
            "osbrowser-smoke-{}",
            uuid::Uuid::new_v4().simple()
        ));
        let child = std::process::Command::new(chrome)
            .arg("--headless=new")
            .arg("--no-first-run")
            .arg("--no-default-browser-check")
            .arg(format!("--user-data-dir={}", path.display()))
            .arg(browser.url())
            .stdout(std::process::Stdio::null())
            .stderr(std::process::Stdio::piped())
            .spawn()
            .map_err(|e| e.to_string())?;
        headless = Some(child);
        profile = Some(path);
    } else if args.iter().any(|arg| arg == "--manual") {
        println!("Open {}", browser.url());
        while !browser.is_connected() {
            std::thread::sleep(std::time::Duration::from_millis(20));
        }
    } else {
        browser.open()?;
    }
    let result = run(&browser, smoke);
    if let Some(mut child) = headless {
        let _ = child.kill();
        if let Ok(output) = child.wait_with_output() {
            if result.is_err() {
                eprintln!("headless Chrome: {}", String::from_utf8_lossy(&output.stderr));
            }
        }
    }
    if let Some(path) = profile {
        let _ = std::fs::remove_dir_all(path);
    }
    result
}

fn run(browser: &osbrowser::BrowserSession, smoke: bool) -> Result<(), String> {
    browser.call(1, "NewDocument", json!({ "title": "OsBrowser" }))?;
    let node = browser.call(1, "CreateElement", json!({
        "tag": "button", "input_type": ""
    }),)?;
    let node = node["Node"].as_u64().ok_or("browser returned no element")?;
    browser.call(1, "SetTextContent", json!([node, "Click me"]))?;
    browser.call(1, "AppendChild", json!({ "parent": 0, "child": node }))?;
    if smoke {
        let html = browser.call(1, "OuterHtml", json!(node))?;
        if !html["Text"].as_str().is_some_and(|html| html.contains("Click me")) {
            return Err(format!("DOM round trip failed: {html}"));
        }
        for offset in 0..128u64 {
            let queued_node = 1_000_000_000_000 + offset;
            browser.enqueue(1, "CreateElement", json!({
                "tag": "span", "input_type": ""
            }), Some(queued_node),)?;
            browser.enqueue(1, "SetTextContent", json!([queued_node, format!("item {offset}")]), None,)?;
            browser.enqueue(1, "AppendChild", json!({
                "parent": 0, "child": queued_node
            }), None,)?;
        }
        let spans = browser.call(1, "QuerySelectorAll", json!("span"))?;
        if spans["Nodes"].as_array().is_none_or(|nodes| nodes.len() != 128) {
            return Err(format!("queued DOM mutations failed: {spans}"));
        }
        let canvas = browser.call(1, "CreateElement", json!({
            "tag": "canvas", "input_type": ""
        }),)?;
        let canvas = canvas["Node"].as_u64().ok_or("browser returned no canvas")?;
        browser.call(1, "AppendChild", json!({ "parent": 0, "child": canvas }))?;
        let target = format!("n{canvas}");
        browser.call(1, "CanvasApply", json!({
            "target": target, "payload": { "SetFillStyleCss": "#ff0000" }
        }),)?;
        browser.call(1, "CanvasApply", json!({
            "target": target, "payload": { "FillRect": [0, 0, 2, 2] }
        }),)?;
        let pixel = browser.call(1, "CanvasQuery", json!({
            "target": target,
            "payload": { "GetImageData": { "sx": 0, "sy": 0, "sw": 1, "sh": 1 } }
        }),)?;
        if pixel["Pixels"]["data"][0] != 255 || pixel["Pixels"]["data"][3] != 255 {
            return Err(format!("canvas round trip failed: {pixel}"));
        }
        let rect = browser.call(1, "BoundingClientRect", json!(node))?;
        let bounds = &rect["Rect"];
        let x = bounds["x"].as_f64().unwrap_or(0.0)
            + bounds["width"].as_f64().unwrap_or(0.0) / 2.0;
        let y = bounds["y"].as_f64().unwrap_or(0.0)
            + bounds["height"].as_f64().unwrap_or(0.0) / 2.0;
        browser.call(1, "DispatchPointer", json!({
            "kind": "click", "client_x": x, "client_y": y, "button": 0
        }),)?;
        let deadline = std::time::Instant::now() + std::time::Duration::from_secs(2);
        while std::time::Instant::now() < deadline {
            if browser.drain_events().into_iter().any(|e| e.node == node && e.kind == "click") {
                println!("DOM and click event round trip passed");
                return Ok(());
            }
            std::thread::sleep(std::time::Duration::from_millis(20));
        }
        return Err("browser click event did not arrive".into());
    }
    println!("The button is visible in {}", browser.url());
    while browser.is_connected() {
        for event in browser.drain_events() {
            if event.node == node && event.kind == "click" {
                println!("Button clicked");
            }
        }
        std::thread::sleep(std::time::Duration::from_millis(20));
    }
    Ok(())
}
