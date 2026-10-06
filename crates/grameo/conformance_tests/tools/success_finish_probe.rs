//! Exact baseline/candidate structural, resource and builder-trace comparisons.
use std::{cell::RefCell, fmt::Write, hint::black_box};
type Shape = (usize, String);
macro_rules! shape {
    ($function:ident,$core:ident) => {
        fn $function(tree: $core::ParseTree<'_, '_>) -> Shape {
            fn visit(p: $core::Pair<'_, '_, '_>, d: usize, s: &mut String) {
                write!(s, "{} {:?} {d};", p.rule_name(), p.as_span()).unwrap();
                for c in p.into_inner() {
                    visit(c, d + 1, s)
                }
            }
            let mut s = String::new();
            for p in tree.pairs() {
                visit(p, 0, &mut s)
            }
            (tree.consumed(), s)
        }
    };
}
shape!(old_shape, before);
shape!(new_shape, after);
#[derive(Default)]
struct Audit {
    active: Vec<usize>,
    log: RefCell<Vec<(u8, usize, usize, usize)>>,
}
macro_rules! builder {
    ($core:ident) => {
        impl $core::builder::Builder for Audit {
            fn checkpoint(&self) -> usize {
                let n = self.active.len();
                self.log.borrow_mut().push((0, n, 0, 0));
                n
            }
            fn rollback(&mut self, n: usize) {
                self.log.borrow_mut().push((1, n, 0, 0));
                self.active.truncate(n)
            }
            fn begin(&mut self, r: usize, s: usize) {
                self.log.borrow_mut().push((2, r, s, 0));
                self.active.push(r)
            }
            fn finish(&mut self, r: usize, s: $core::Span, _: &str) {
                self.log.borrow_mut().push((3, r, s.start, s.end));
                self.active.push(r)
            }
        }
    };
}
builder!(before);
builder!(after);
unsafe extern "C" {
    fn js_thread_cpu_ns() -> u64;
}
fn ns() -> u64 {
    let n = unsafe { js_thread_cpu_ns() };
    assert_ne!(n, u64::MAX);
    n
}
fn compare() {
    let source = r#"WHITESPACE = _{ " " | "\t" }
item = { "a" | "b" }
main = { (item ~ ","?)* ~ EOI }
bounded = { (item ~ ","?){1,3} ~ EOI }
predicate = { &(item ~ item) ~ item ~ item ~ EOI }
stack = { PUSH("a") ~ ("b" | "a") ~ POP ~ EOI }
atomic = @{ ("a" ~ "b"?)* ~ EOI }
compound = ${ (item ~ ","?)* ~ EOI }
deep = { ((("a" ~ "b") ~ "c") ~ "d") ~ EOI }
rollback = { (("a" ~ "b") ~ "c") | ("a" ~ "d") }
nullable = { ("" ~ "") ~ EOI }
progress = { ("" ~ "")* ~ EOI }
"#;
    let old = before::resolve(before::parse(source).unwrap()).unwrap();
    let new = after::resolve(after::parse(source).unwrap()).unwrap();
    let mut count = 0;
    for entry in [
        "main",
        "bounded",
        "predicate",
        "stack",
        "atomic",
        "compound",
        "deep",
        "rollback",
        "nullable",
        "progress",
    ] {
        for input in [
            "", "a", "b", "z", "ab", "a,b", "a, b", "a b a b", "a,", " a ", "abcd", "ad", "é",
            "a😀", "a\0",
        ] {
            for depth in [0, 1, 2, 3, 5, 9] {
                for steps in 0..=160 {
                    let a = before::MatchOptions {
                        max_steps: steps,
                        max_rule_depth: depth,
                    };
                    let b = after::MatchOptions {
                        max_steps: steps,
                        max_rule_depth: depth,
                    };
                    let x = old
                        .parse_source_with_options(entry, input, a)
                        .map(old_shape)
                        .map_err(|e| format!("{e:?}"));
                    let y = new
                        .parse_source_with_options(entry, input, b)
                        .map(new_shape)
                        .map_err(|e| format!("{e:?}"));
                    assert_eq!(x, y, "captures {entry} {input:?} {steps} {depth}");
                    assert_eq!(
                        format!("{:?}", old.recognize_with_options(entry, input, a)),
                        format!("{:?}", new.recognize_with_options(entry, input, b)),
                        "recognition {entry} {input:?} {steps} {depth}"
                    );
                    let (mut aa, mut bb) = (Audit::default(), Audit::default());
                    let x = before::engine::build_program(&old, entry, input, a, &mut aa);
                    let y = after::engine::build_program(&new, entry, input, b, &mut bb);
                    assert_eq!(format!("{x:?}"), format!("{y:?}"));
                    assert_eq!(aa.active, bb.active);
                    assert_eq!(
                        aa.log.into_inner(),
                        bb.log.into_inner(),
                        "builder {entry} {input:?} {steps} {depth}"
                    );
                    count += 1;
                }
            }
        }
    }
    println!("PARITY {count} captures/errors/recognition logical costs/resources/builder traces");
}
fn screen() {
    let grammar = r#"root = { ("a" ~ "b" ~ "c" ~ "d" ~ "e" ~ "f" ~ "g" ~ "h")* ~ EOI }"#;
    let old = before::compile(grammar).unwrap();
    let new = after::compile(grammar).unwrap();
    for n in [1, 16, 1024] {
        let input = "abcdefgh".repeat(n);
        assert_eq!(
            old_shape(old.parse_source("root", &input).unwrap()),
            new_shape(new.parse_source("root", &input).unwrap())
        );
        let mut samples = [vec![], vec![]];
        for round in 0..111 {
            for i in 0..2 {
                let variant = (i + round) % 2;
                let start = ns();
                for _ in 0..16 {
                    if variant == 0 {
                        drop(black_box(old.parse_source("root", black_box(&input))));
                    } else {
                        drop(black_box(new.parse_source("root", black_box(&input))));
                    }
                }
                let elapsed = (ns() - start) as f64 / 16.;
                if round >= 10 {
                    samples[variant].push(elapsed);
                }
            }
        }
        for s in &mut samples {
            s.sort_by(f64::total_cmp);
        }
        println!(
            "SCREEN {n} {} {} {}",
            samples[0][50],
            samples[1][50],
            samples[1][50] / samples[0][50]
        );
    }
}
fn main() {
    std::thread::Builder::new()
        .stack_size(64 * 1024 * 1024)
        .spawn(|| {
            if let Some(manifest) = std::env::args().nth(1) {
                screen_real(&manifest);
            } else {
                compare();
                screen();
            }
        })
        .unwrap()
        .join()
        .unwrap();
}

fn screen_real(manifest: &str) {
    for line in std::fs::read_to_string(manifest).unwrap().lines() {
        let fields = line.split('\t').collect::<Vec<_>>();
        let source = std::fs::read_to_string(fields[2]).unwrap();
        let grammar = std::fs::read_to_string(fields[3]).unwrap();
        let old = before::compile(&grammar).unwrap();
        let new = after::compile(&grammar).unwrap();
        let entry = if fields[0] == "lua" {
            "chunk"
        } else {
            "program"
        };
        let a = old
            .parse_source(entry, &source)
            .map(old_shape)
            .map_err(|e| format!("{e:?}"));
        let b = new
            .parse_source(entry, &source)
            .map(new_shape)
            .map_err(|e| format!("{e:?}"));
        assert_eq!(a, b);
        assert_eq!(
            format!("{:?}", old.recognize(entry, &source)),
            format!("{:?}", new.recognize(entry, &source))
        );
        let n = 8;
        let mut times = [vec![], vec![]];
        for round in 0..511 {
            for j in 0..2 {
                let backend = (round + j) % 2;
                let start = ns();
                for _ in 0..n {
                    if backend == 0 {
                        drop(black_box(old.parse_source(entry, black_box(&source))));
                    } else {
                        drop(black_box(new.parse_source(entry, black_box(&source))));
                    }
                }
                let elapsed = (ns() - start) as f64 / n as f64;
                if round >= 10 {
                    times[backend].push(elapsed);
                }
            }
        }
        for v in &mut times {
            v.sort_by(f64::total_cmp);
        }
        println!(
            "REAL\t{}\t{}\t{}\t{}\t{}\t{}\t{}",
            fields[0],
            fields[1],
            times[0][250],
            times[1][250],
            times[1][250] / times[0][250],
            times[0][475],
            times[1][475]
        );
    }
}
