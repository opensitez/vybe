//! Complete source-to-vybe_ast throughput, using Cargo's actual Python library.
//! No emission, linking or VM execution is measured. Run without --release.
use std::{error::Error, hint::black_box, path::PathBuf, time::Instant};

#[cfg(all(
    target_pointer_width = "64",
    any(target_os = "macos", target_os = "linux")
))]
fn process_cpu() -> std::io::Result<f64> {
    #[repr(C)]
    struct Timespec {
        seconds: i64,
        nanoseconds: i64,
    }
    unsafe extern "C" {
        fn clock_gettime(clock: i32, time: *mut Timespec) -> i32;
    }
    #[cfg(target_os = "macos")]
    const PROCESS_CPU_CLOCK: i32 = 12;
    #[cfg(target_os = "linux")]
    const PROCESS_CPU_CLOCK: i32 = 2;
    let mut time = Timespec {
        seconds: 0,
        nanoseconds: 0,
    };
    if unsafe { clock_gettime(PROCESS_CPU_CLOCK, &mut time) } != 0 {
        return Err(std::io::Error::last_os_error());
    }
    Ok(time.seconds as f64 + time.nanoseconds as f64 * 1e-9)
}

#[cfg(not(all(
    target_pointer_width = "64",
    any(target_os = "macos", target_os = "linux")
)))]
fn process_cpu() -> std::io::Result<f64> {
    Err(std::io::Error::new(
        std::io::ErrorKind::Unsupported,
        "this benchmark needs a supported process CPU clock",
    ))
}

fn main() -> Result<(), Box<dyn Error>> {
    if !cfg!(debug_assertions) {
        return Err("use the DEBUG build for this benchmark (no --release)".into());
    }
    let mut iterations = 6usize;
    let mut include_rejects = false;
    let mut paths = Vec::new();
    let mut arguments = std::env::args().skip(1);
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--iterations" => {
                iterations = arguments.next().ok_or("missing iteration count")?.parse()?
            }
            "--include-rejects" => include_rejects = true,
            "--help" => {
                println!(
                    "cargo run -p vybe_language_python --example parse_throughput -- [--iterations N] [--include-rejects] file.py ..."
                );
                return Ok(());
            }
            option if option.starts_with('-') => {
                return Err(format!("unknown option: {option}").into());
            }
            _ => paths.push(PathBuf::from(argument)),
        }
    }
    if iterations == 0 || paths.is_empty() {
        return Err("supply a positive iteration count and at least one Python source path".into());
    }
    vybe_language_python::register();
    type Parser = fn(&str) -> Result<vybe_ast::Module, String>;
    let parsers: [(&str, Parser); 2] = [
        ("registered_hybrid", vybe_language_python::parse),
        ("peg", vybe_language_python::parse_with_peg),
    ];
    for path in paths {
        let source = std::fs::read_to_string(&path)?;
        let physical_lines = source.bytes().filter(|byte| *byte == b'\n').count()
            + usize::from(!source.is_empty() && !source.ends_with('\n'));
        // Full output/error parity and one-time initialization are outside samples.
        let original = parsers[0].1(&source);
        let reference = parsers[1].1(&source);
        if format!("{original:?}") != format!("{reference:?}") {
            return Err(format!("complete Module/error PEG mismatch: {}", path.display()).into());
        }
        let accepted = original.is_ok();
        let error = original.as_ref().err().cloned();
        drop((original, reference));
        if !accepted && !include_rejects {
            return Err(format!("{} rejected: {}", path.display(), error.unwrap()).into());
        }
        let mut samples = [Vec::new(), Vec::new()];
        for iteration in 0..iterations {
            // Alternate modes to reduce systematic first/second bias.
            for index in if iteration % 2 == 0 { [0, 1] } else { [1, 0] } {
                let cpu = process_cpu()?;
                let wall = Instant::now();
                let result = parsers[index].1(black_box(&source));
                let parse_wall = wall.elapsed().as_secs_f64();
                let parse_cpu = process_cpu()? - cpu;
                if result.is_ok() != accepted {
                    return Err(format!("parse outcome changed: {}", path.display()).into());
                }
                black_box(&result);
                let cpu = process_cpu()?;
                drop(black_box(result));
                samples[index].push((parse_cpu, parse_wall, process_cpu()? - cpu));
            }
        }
        for (index, (mode, _)) in parsers.into_iter().enumerate() {
            let parse_cpu: f64 = samples[index].iter().map(|sample| sample.0).sum();
            let drop_cpu: f64 = samples[index].iter().map(|sample| sample.2).sum();
            if parse_cpu <= 0.0 {
                return Err("CPU clock did not advance; increase iterations".into());
            }
            // Rejected input timing is diagnostic; it is never accepted-AST throughput.
            let bytes_per_cpu_second =
                accepted.then(|| source.len() as f64 * iterations as f64 / parse_cpu);
            let lines_per_cpu_second =
                accepted.then(|| physical_lines as f64 * iterations as f64 / parse_cpu);
            let lifecycle_lines_per_cpu_second = accepted
                .then(|| physical_lines as f64 * iterations as f64 / (parse_cpu + drop_cpu));
            println!(
                "{}",
                serde_json::json!({
                    "mode": mode, "file": path, "accepted": accepted, "error": error,
                    "bytes": source.len(), "physical_lines": physical_lines, "iterations": iterations,
                    "complete_module_parse_cpu_seconds": parse_cpu, "ast_drop_cpu_seconds": drop_cpu,
                    "bytes_per_parse_cpu_second": bytes_per_cpu_second,
                    "physical_lines_per_parse_cpu_second": lines_per_cpu_second,
                    "physical_lines_per_parse_and_drop_cpu_second": lifecycle_lines_per_cpu_second,
                    "samples_parse_cpu_wall_drop_cpu": samples[index],
                    "scope": "actual public parser to complete Module; required hooks, scopes and normalization included; drop reported separately; file I/O, emission, linking and VM execution excluded",
                    "physical_line_definition": "LF count plus unterminated final line; includes blank/comment lines",
                    "clock": "process CPU", "debug_assertions": cfg!(debug_assertions),
                    "binary": std::env::current_exe()?,
                })
            );
        }
    }
    Ok(())
}
