//! Matched complete-frontend CPU comparison, excluding compilation and VM work.
//! cargo run -p vybe_language_python --example frontend_cpu -- file.py 25
use std::{hint::black_box, time::Instant};

#[cfg(any(target_os = "macos", target_os = "linux"))]
fn cpu_seconds() -> f64 {
    #[repr(C)]
    struct Timespec { seconds: i64, nanoseconds: i64 }
    unsafe extern "C" {
        fn clock_gettime(clock: i32, time: *mut Timespec) -> i32;
    }
    #[cfg(target_os = "macos")]
    const PROCESS_CPU_CLOCK: i32 = 12;
    #[cfg(target_os = "linux")]
    const PROCESS_CPU_CLOCK: i32 = 2;
    let mut time = Timespec { seconds: 0, nanoseconds: 0 };
    assert_eq!(unsafe { clock_gettime(PROCESS_CPU_CLOCK, &mut time) }, 0);
    time.seconds as f64 + time.nanoseconds as f64 * 1e-9
}

#[cfg(not(any(target_os = "macos", target_os = "linux")))]
fn cpu_seconds() -> f64 { f64::NAN }

fn main() {
    let args: Vec<_> = std::env::args().collect();
    let path = args.get(1).expect("pass a Python source path");
    let iterations: usize = args.get(2).map_or(25, |arg| arg.parse().unwrap());
    assert!(iterations > 0);
    let source = std::fs::read_to_string(path).unwrap();
    type Frontend = fn(&str) -> Result<vybe_ast::Module, String>;
    let frontends: [(&str, Frontend); 2] = [
        ("peg", vybe_language_python::parse_with_peg),
        ("registered_hybrid", vybe_language_python::parse),
    ];
    for (_, parse) in frontends {
        for _ in 0..3 { black_box(parse(&source)).unwrap(); }
    }
    for sample in 0..10 {
        for index in if sample % 2 == 0 { [0, 1] } else { [1, 0] } {
            let (mode, parse) = frontends[index];
            let cpu = cpu_seconds();
            let wall = Instant::now();
            for _ in 0..iterations { black_box(parse(&source)).unwrap(); }
            println!("sample={sample} mode={mode} iterations={iterations} cpu_seconds={:.6} wall_seconds={:.6}",
                cpu_seconds() - cpu, wall.elapsed().as_secs_f64());
        }
    }
}
