//! Complete source-to-vybe_ast throughput; no emission, linking or execution.
use std::{error::Error, hint::black_box, path::PathBuf, time::Instant};

#[cfg(unix)]
fn thread_cpu() -> std::io::Result<f64> {
    let mut time = libc::timespec {
        tv_sec: 0,
        tv_nsec: 0,
    };
    if unsafe { libc::clock_gettime(libc::CLOCK_THREAD_CPUTIME_ID, &mut time) } != 0 {
        return Err(std::io::Error::last_os_error());
    }
    Ok(time.tv_sec as f64 + time.tv_nsec as f64 * 1e-9)
}

#[cfg(not(unix))]
fn thread_cpu() -> std::io::Result<f64> {
    Err(std::io::Error::new(
        std::io::ErrorKind::Unsupported,
        "this benchmark needs a thread CPU clock",
    ))
}

fn preparation_parser(source: &str) -> Result<vybe_ast::Module, String> {
    vybe_language_php::parse_with_source_unit_preparation(source, &mut |requests| {
        black_box(requests);
    })
}

fn main() -> Result<(), Box<dyn Error>> {
    if !cfg!(debug_assertions) {
        return Err("use the DEBUG build for this benchmark".into());
    }
    let mut iterations = 3usize;
    let mut paths = Vec::new();
    let mut arguments = std::env::args().skip(1);
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--iterations" => {
                iterations = arguments.next().ok_or("missing iteration count")?.parse()?
            }
            "--help" => {
                println!("cargo run -p vybe_language_php --example parse_throughput -- [--iterations N] [file.php ...]");
                return Ok(());
            }
            option if option.starts_with('-') => {
                return Err(format!("unknown option: {option}").into())
            }
            _ => paths.push(PathBuf::from(argument)),
        }
    }
    if iterations == 0 {
        return Err("iterations must be positive".into());
    }
    if paths.is_empty() {
        let root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../examples/webroot/wordpress/wp-includes");
        paths.extend(
            [
                "functions.php",
                "class-wp-hook.php",
                "formatting.php",
                "script-loader.php",
            ]
            .map(|name| root.join(name)),
        );
    }
    vybe_language_php::register();
    type ParseFn = fn(&str) -> Result<vybe_ast::Module, String>;
    let parsers: [(&str, ParseFn); 3] = [
        ("optimized", vybe_language_php::parse),
        ("peg", vybe_language_php::parse_with_peg),
        ("production_preparation", preparation_parser),
    ];
    for path in paths {
        let source = std::fs::read_to_string(&path)?;
        let lines = source.lines().count();
        // Check the complete reference result and warm immutable process state
        // outside the timed samples. A divergent AST earns no throughput row.
        let optimized = parsers[0].1(&source)?;
        let reference = parsers[1].1(&source)?;
        let prepared = parsers[2].1(&source)?;
        if format!("{optimized:?}") != format!("{reference:?}")
            || format!("{prepared:?}") != format!("{reference:?}")
        {
            return Err(format!("optimized/PEG AST mismatch: {}", path.display()).into());
        }
        drop((optimized, reference, prepared));
        let mut samples = [Vec::new(), Vec::new(), Vec::new()];
        for iteration in 0..iterations {
            // Rotate all three paths through each position.
            for position in 0..parsers.len() {
                let index = (position + iteration) % parsers.len();
                let (_, parser) = parsers[index];
                let cpu = thread_cpu()?;
                let wall = Instant::now();
                let module = parser(black_box(&source))?;
                let parse_wall = wall.elapsed().as_secs_f64();
                let parse_cpu = thread_cpu()? - cpu;
                black_box(&module);
                let cpu = thread_cpu()?;
                drop(black_box(module));
                samples[index].push((parse_cpu, parse_wall, thread_cpu()? - cpu));
            }
        }
        for (index, (mode, _)) in parsers.into_iter().enumerate() {
            let parse_cpu: f64 = samples[index].iter().map(|sample| sample.0).sum();
            println!(
                "{}",
                serde_json::json!({
                    "mode": mode, "file": path, "bytes": source.len(), "physical_lines": lines,
                    "iterations": iterations, "parse_cpu_seconds": parse_cpu,
                    "bytes_per_cpu_second": source.len() as f64 * iterations as f64 / parse_cpu,
                    "lines_per_cpu_second": lines as f64 * iterations as f64 / parse_cpu,
                    "samples_parse_cpu_wall_drop_cpu": samples[index],
                    "scope": "public parser to complete Module; required hooks and AST construction included; AST drop reported separately; file I/O, emission, linking and VM execution excluded",
                    "debug_assertions": cfg!(debug_assertions), "binary": std::env::current_exe()?,
                "complete_debug_ast_equality_checked": true,
                })
            );
        }
    }
    Ok(())
}
