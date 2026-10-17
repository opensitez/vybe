//! Complete public-parser/PEG Module and error equivalence, one input at a time.
//! No timing, emission or VM execution. Temporary Debug output is deleted per input.
#[path = "support/exact_debug.rs"]
mod exact_debug;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let mut paths = Vec::new();
    while let Some(arg) = args.next() {
        if arg == "--manifest" {
            let manifest = args
                .next()
                .ok_or("missing newline-delimited manifest path")?;
            paths.extend(
                std::fs::read_to_string(manifest)?
                    .lines()
                    .filter(|line| !line.is_empty())
                    .map(str::to_owned),
            );
        } else if arg.starts_with('-') {
            return Err(format!("unknown option: {arg}").into());
        } else {
            paths.push(arg);
        }
    }
    if paths.is_empty() {
        return Err("supply Python paths or --manifest PATH".into());
    }
    vybe_language_python::register();
    let (mut accepted, mut rejected, mut input_errors) = (0, 0, 0);
    for path in paths {
        let source = match std::fs::read_to_string(&path) {
            Ok(source) => source,
            Err(error) if error.kind() == std::io::ErrorKind::InvalidData => {
                input_errors += 1;
                eprintln!("input-encoding-error path={path:?}: {error}");
                continue;
            }
            Err(error) => return Err(error.into()),
        };
        let actual = vybe_language_python::parse(&source);
        if actual.is_ok() {
            accepted += 1;
        } else {
            rejected += 1;
        }
        let snapshot = exact_debug::Snapshot::new(&actual)?;
        drop(actual);
        let reference = vybe_language_python::parse_with_peg(&source);
        if let Some(difference) = snapshot.compare(&reference)? {
            return Err(
                format!("complete Module/error mismatch path={path:?}: {difference:?}").into(),
            );
        }
    }
    println!(
        "exact public/PEG accepted={accepted} rejected={rejected} input_encoding_errors={input_errors}"
    );
    Ok(())
}
