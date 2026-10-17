//! Measure the browser transport directly, without the compiler or WASM VM.
use bytes::Bytes;
use osbrowser::BrowserSession;
use serde_json::{Value, json};
use std::time::{Duration, Instant};

fn canvas(
    browser: &BrowserSession,
    width: u32,
    height: u32,
    visible: bool,
) -> Result<String, String> {
    let result = browser.call(
        1,
        "CreateElement",
        json!({"tag": "canvas", "input_type": ""}),
    )?;
    let node = result["Node"].as_u64().ok_or("missing canvas node")?;
    for (name, value) in [("width", width), ("height", height)] {
        browser.call(1, "SetAttribute", json!([node, name, value.to_string()]))?;
    }
    if visible {
        browser.call(1, "AppendChild", json!({"parent": 0, "child": node}))?;
    }
    let target = format!("n{node}");
    browser.call(1, "CanvasEnsure", json!({"target": target}))?;
    browser.call(
        1,
        "CanvasApply",
        json!({"target": target,
        "payload": {"SetImageSmoothing": false}}),
    )?;
    browser.call(
        1,
        "CanvasApply",
        json!({"target": target,
        "payload": {"SetGlobalCompositeOperation": "copy"}}),
    )?;
    Ok(target)
}

fn run(
    browser: &BrowserSession,
    frames: usize,
    queued: bool,
    copy_only: bool,
    binary: bool,
) -> Result<(), String> {
    browser.call(
        1,
        "NewDocument",
        json!({"title": "Direct Rust Canvas Bridge Benchmark"}),
    )?;
    let source = canvas(browser, 320, 200, false)?;
    let back = canvas(browser, 960, 600, false)?;
    let front = canvas(browser, 960, 600, true)?;
    let mut images = Vec::new();
    let mut raw_images = Vec::new();
    for phase in 0..8u32 {
        let mut pixels = Vec::with_capacity(320 * 200 * 4);
        for y in 0..200u32 {
            for x in 0..320u32 {
                let bright = ((x + phase * 8) / 16 + y / 16) % 2;
                pixels.extend_from_slice(&[
                    (32 + bright * 144) as u8,
                    (48 + phase * 24) as u8,
                    (80 + bright * 120) as u8,
                    255,
                ]);
            }
        }
        let mut image = json!({"target": source, "payload": {"PutImageData": {
            "width": 320, "height": 200, "dx": 0, "dy": 0,
        }}});
        if !binary {
            image["payload"]["PutImageData"]["pixels"] = json!(pixels);
        }
        raw_images.push(Bytes::from(pixels));
        images.push(image);
    }
    let copies = [
        json!({"target": back, "payload": {"source": source,
            "rect": [0, 0, 320, 200, 0, 0, 960, 600]}}),
        json!({"target": front, "payload": {"source": back,
            "rect": [0, 0, 960, 600, 0, 0, 960, 600]}}),
    ];
    let mut cached_sources = Vec::new();
    if copy_only {
        for (index, image) in images.iter().enumerate() {
            let target = canvas(browser, 320, 200, false)?;
            let mut image = image.clone();
            image["target"] = json!(target);
            if binary {
                browser.call_with_pixels(1, image, raw_images[index].clone())?;
            } else {
                browser.call(1, "CanvasApply", image)?;
            }
            cached_sources.push(target);
        }
    }
    let ping = Instant::now();
    for _ in 0..32 {
        browser.call(1, "Title", Value::Null)?;
    }
    println!(
        "empty DOM roundtrip: {:.3} ms average",
        ping.elapsed().as_secs_f64() * 1000.0 / 32.0
    );

    let mut preparation = Duration::ZERO;
    let mut upload = Duration::ZERO;
    let mut copy = Duration::ZERO;
    let mut barrier = Duration::ZERO;
    let mut maximum = Duration::ZERO;
    let start = Instant::now();
    for index in 0..frames {
        let frame_start = Instant::now();
        let image = (!copy_only).then(|| images[index % images.len()].clone());
        let mut first_copy = copies[0].clone();
        if copy_only {
            first_copy["payload"]["source"] = json!(cached_sources[index % cached_sources.len()]);
        }
        preparation += frame_start.elapsed();
        if queued {
            browser.begin_batch();
        }
        let a = Instant::now();
        if let Some(image) = image {
            if binary && queued {
                browser.enqueue_with_pixels(
                    1,
                    image,
                    raw_images[index % raw_images.len()].clone(),
                )?;
            } else if binary {
                browser.call_with_pixels(1, image, raw_images[index % raw_images.len()].clone())?;
            } else if queued {
                browser.enqueue(1, "CanvasApply", image, None)?;
            } else {
                browser.call(1, "CanvasApply", image)?;
            }
        }
        upload += a.elapsed();
        let b = Instant::now();
        for args in [&first_copy, &copies[1]] {
            if queued {
                browser.enqueue(1, "CanvasDrawImage", args.clone(), None)?;
            } else {
                browser.call(1, "CanvasDrawImage", args.clone())?;
            }
        }
        copy += b.elapsed();
        if queued {
            browser.end_batch();
            let c = Instant::now();
            // The next synchronous call observes all preceding mutations.
            browser.call(1, "Title", Value::Null)?;
            barrier += c.elapsed();
        }
        maximum = maximum.max(frame_start.elapsed());
    }
    let elapsed = start.elapsed();
    let ms = |duration: Duration| duration.as_secs_f64() * 1000.0 / frames as f64;
    println!(
        "mode={} content={} frames={} elapsed_ms={:.1} completed_submissions_fps={:.2}",
        if queued {
            "queued-with-barrier"
        } else {
            "synchronous"
        },
        if copy_only {
            "browser-cached"
        } else {
            "pixel-upload"
        },
        frames,
        elapsed.as_secs_f64() * 1000.0,
        frames as f64 / elapsed.as_secs_f64()
    );
    println!(
        "pixel_transport={}",
        if binary {
            "binary-rgba"
        } else {
            "json-numbers"
        }
    );
    println!(
        "avg_ms: prepare={:.3} upload={:.3} two_canvas_copies={:.3} barrier={:.3}; max_frame_ms={:.3}",
        ms(preparation),
        ms(upload),
        ms(copy),
        ms(barrier),
        maximum.as_secs_f64() * 1000.0
    );
    let pixel = browser.call(
        1,
        "CanvasQuery",
        json!({"target": front,
        "payload": {"GetImageData": {"sx": 0, "sy": 0, "sw": 1, "sh": 1}}}),
    )?;
    let phase = ((frames - 1) % 8) as u32;
    let bright = (phase * 8 / 16) % 2;
    let expected = json!([32 + bright * 144, 48 + phase * 24, 80 + bright * 120, 255]);
    if pixel["Pixels"]["data"] != expected {
        return Err(format!("pixel verification failed: {pixel}"));
    }
    println!("final destination pixel verified; no VM or C compiler used");
    Ok(())
}

fn main() -> Result<(), String> {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let mut frames = 120;
    let mut queued = false;
    let mut copy_only = false;
    let mut headless = false;
    let mut manual = false;
    let mut application = None;
    let mut binary = false;
    let mut i = 0;
    while i < args.len() {
        match args[i].as_str() {
            "--frames" => {
                i += 1;
                frames = args
                    .get(i)
                    .ok_or("missing frame count")?
                    .parse::<usize>()
                    .map_err(|e| e.to_string())?;
                if frames == 0 {
                    return Err("frame count must be positive".into());
                }
            }
            "--queued" => queued = true,
            "--copy-only" => copy_only = true,
            "--binary" => binary = true,
            "--headless" => headless = true,
            "--manual" => manual = true,
            "--browser" => {
                i += 1;
                application = Some(match args.get(i).map(String::as_str) {
                    Some("chrome") => "Google Chrome",
                    Some("safari") => "Safari",
                    Some("firefox") => "Firefox",
                    _ => return Err("choose --browser chrome, safari, or firefox".into()),
                });
            }
            _ => {
                return Err(
                    "Usage: canvas_bridge_bench [--frames N] [--binary] [--queued] [--copy-only] [--headless | --manual | --browser chrome|safari|firefox]"
                        .into(),
                );
            }
        }
        i += 1;
    }
    if (manual && headless) || (application.is_some() && (manual || headless)) {
        return Err("choose one of --manual, --headless, or --browser".into());
    }
    let browser = BrowserSession::start()?;
    let mut child = None;
    let mut profile = None;
    if headless {
        #[cfg(target_os = "macos")]
        let chrome = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
        #[cfg(target_os = "windows")]
        let chrome = "chrome.exe";
        #[cfg(all(unix, not(target_os = "macos")))]
        let chrome = "google-chrome";
        let path = std::env::temp_dir().join(format!(
            "canvas-bridge-bench-{}",
            uuid::Uuid::new_v4().simple()
        ));
        child = Some(
            std::process::Command::new(chrome)
                .args([
                    "--headless=new",
                    "--no-first-run",
                    "--no-default-browser-check",
                ])
                .arg(format!("--user-data-dir={}", path.display()))
                .arg(browser.url())
                .stdout(std::process::Stdio::null())
                .stderr(std::process::Stdio::null())
                .spawn()
                .map_err(|e| e.to_string())?,
        );
        profile = Some(path);
    } else if manual {
        println!("Open {} in the browser to measure", browser.url());
    } else if let Some(application) = application {
        #[cfg(target_os = "macos")]
        {
            let status = std::process::Command::new("open")
                .args(["-a", application, &browser.url()])
                .status()
                .map_err(|e| e.to_string())?;
            if !status.success() {
                return Err(format!("could not open {application}: {status}"));
            }
            println!("browser={application}");
        }
        #[cfg(not(target_os = "macos"))]
        return Err(format!(
            "--browser selection is only implemented on macOS; use --manual for {application}"
        ));
    } else {
        browser.open()?;
    }
    let result = (|| {
        let deadline = Instant::now() + Duration::from_secs(120);
        while !browser.is_connected() {
            if Instant::now() >= deadline {
                return Err("browser did not connect".into());
            }
            std::thread::sleep(Duration::from_millis(20));
        }
        run(&browser, frames, queued, copy_only, binary)
    })();
    if let Some(mut child) = child {
        let _ = child.kill();
        let _ = child.wait();
    }
    if let Some(path) = profile {
        let _ = std::fs::remove_dir_all(path);
    }
    result
}
