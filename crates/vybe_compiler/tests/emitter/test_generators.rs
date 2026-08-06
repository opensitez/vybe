use std::fs;
use std::path::PathBuf;

#[test]
fn generator_stack_switching_ops_are_emitted_only_by_generator_emitter() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("src");
    let allowed = root.join("primitives").join("generators.rs");
    let needles = [
        "Op::SUSPEND",
        "Op::RESUME",
        "Op::RESUME_THROW",
        "Op::GEN_NEXT",
    ];

    let mut offenders = Vec::new();
    visit_rs_files(&root, &mut |path, source| {
        if *path == allowed {
            return;
        }
        for (line_number, line) in source.lines().enumerate() {
            let trimmed = line.trim_start();
            if trimmed.starts_with("//") || trimmed.starts_with("///") || trimmed.starts_with("//!")
            {
                continue;
            }
            for needle in needles {
                if !line.contains(needle) {
                    continue;
                }
                offenders.push(format!(
                    "{}:{} contains {}",
                    path.strip_prefix(&root).unwrap_or(path).display(),
                    line_number + 1,
                    needle
                ));
            }
        }
    });

    assert!(
        offenders.is_empty(),
        "generator stack-switching opcodes must go through emitter::generators:\n{}",
        offenders.join("\n")
    );
}

#[test]
fn promise_suspend_ops_are_emitted_only_by_common_async_emitter() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("src");
    let allowed = root.join("primitives").join("functions.rs");
    let needle = "Op::PROMISE_SUSPEND";

    let mut offenders = Vec::new();
    visit_rs_files(&root, &mut |path, source| {
        if *path == allowed {
            return;
        }
        for (line_number, line) in source.lines().enumerate() {
            let trimmed = line.trim_start();
            if trimmed.starts_with("//") || trimmed.starts_with("///") || trimmed.starts_with("//!")
            {
                continue;
            }
            if line.contains(needle) {
                offenders.push(format!(
                    "{}:{} contains {}",
                    path.strip_prefix(&root).unwrap_or(path).display(),
                    line_number + 1,
                    needle
                ));
            }
        }
    });

    assert!(
        offenders.is_empty(),
        "promise suspend opcodes must go through emitter::functions::emit_await:\n{}",
        offenders.join("\n")
    );
}

fn visit_rs_files(dir: &PathBuf, f: &mut impl FnMut(&PathBuf, &str)) {
    for entry in fs::read_dir(dir).expect("read source dir") {
        let entry = entry.expect("read dir entry");
        let path = entry.path();
        if path.is_dir() {
            visit_rs_files(&path, f);
        } else if path.extension().is_some_and(|ext| ext == "rs") {
            let source = fs::read_to_string(&path).expect("read rust source");
            f(&path, &source);
        }
    }
}
