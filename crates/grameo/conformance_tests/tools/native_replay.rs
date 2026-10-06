//! Pest-free full capture/recognition replay against one dependency closure.
use std::{
    env, fs,
    io::{self, BufReader, BufWriter, Read, Write},
};
use vybe_parser::{MatchOptions, ParseError, program::Program};
mod python_helper {
    include!(env!("VYBE_AUDIT_PYTHON_HELPER"));
}
#[allow(clippy::collapsible_if)]
mod fortran_helper {
    include!(env!("VYBE_AUDIT_FORTRAN_HELPER"));
}

fn number(reader: &mut impl Read) -> u64 {
    let mut bytes = [0; 8];
    reader.read_exact(&mut bytes).unwrap();
    u64::from_le_bytes(bytes)
}
fn string(reader: &mut impl Read) -> String {
    let length = number(reader) as usize;
    let mut bytes = vec![0; length];
    reader.read_exact(&mut bytes).unwrap();
    String::from_utf8(bytes).unwrap()
}
fn push_number(out: &mut Vec<u8>, value: usize) {
    out.extend_from_slice(&(value as u64).to_le_bytes());
}
fn push_string(out: &mut Vec<u8>, value: &str) {
    push_number(out, value.len());
    out.extend_from_slice(value.as_bytes());
}
fn fingerprint(label: &str, source: &str) -> u64 {
    let mut value = 0xcbf29ce484222325u64;
    for part in [label.as_bytes(), source.as_bytes()] {
        for byte in (part.len() as u64).to_le_bytes().iter().chain(part) {
            value = (value ^ u64::from(*byte)).wrapping_mul(0x100000001b3);
        }
    }
    value
}
fn error(out: &mut Vec<u8>, value: &ParseError) {
    out.push(0);
    push_string(out, &format!("{:?}", value.kind));
    push_string(out, &format!("{value:?}"));
}
fn captures<P: Program>(
    program: &P,
    entry: &str,
    source: &str,
    options: MatchOptions,
    expected: bool,
) -> Vec<u8> {
    let mut out = Vec::new();
    match vybe_parser::engine::parse_program(program, entry, source, options) {
        Ok(tree) => {
            let _ = expected;
            out.push(1);
            push_number(&mut out, tree.consumed());
            let mut nodes = Vec::new();
            fn visit(
                pair: vybe_parser::Pair<'_, '_, '_>,
                depth: usize,
                nodes: &mut Vec<(String, usize, usize, usize)>,
            ) {
                let span = pair.as_span();
                nodes.push((pair.rule_name().to_owned(), span.start, span.end, depth));
                for child in pair.into_inner() {
                    visit(child, depth + 1, nodes);
                }
            }
            for pair in tree.pairs() {
                visit(pair, 0, &mut nodes);
            }
            push_number(&mut out, nodes.len());
            for (name, start, end, depth) in nodes {
                push_string(&mut out, &name);
                push_number(&mut out, start);
                push_number(&mut out, end);
                push_number(&mut out, depth);
            }
        }
        Err(value) => {
            error(&mut out, &value);
        }
    }
    out
}
fn recognition<P: Program>(
    program: &P,
    entry: &str,
    source: &str,
    options: MatchOptions,
    expected: bool,
) -> Vec<u8> {
    let mut out = Vec::new();
    match vybe_parser::engine::recognize_program(program, entry, source, options) {
        Ok(value) => {
            let _ = expected;
            out.push(1);
            push_number(&mut out, value.consumed);
            push_number(&mut out, value.steps);
        }
        Err(value) => {
            error(&mut out, &value);
        }
    }
    out
}
fn case<P: Program>(
    resolved: &vybe_parser::CompiledGrammar,
    generated: &P,
    entry: &str,
    label: &str,
    source: &str,
    expected: (bool, usize, u64),
    options: MatchOptions,
    writer: &mut impl Write,
) {
    assert_eq!(source.len(), expected.1, "{label}: bytes");
    assert_eq!(
        fingerprint(label, source),
        expected.2,
        "{label}: source fingerprint"
    );
    let mut record = Vec::new();
    push_number(&mut record, source.len());
    record.extend_from_slice(&expected.2.to_le_bytes());
    record.push(u8::from(expected.0));
    let outcomes = [
        captures(resolved, entry, source, options, expected.0),
        captures(generated, entry, source, options, expected.0),
        recognition(resolved, entry, source, options, expected.0),
        recognition(generated, entry, source, options, expected.0),
    ];
    for outcome in outcomes {
        push_number(&mut record, outcome.len());
        record.extend_from_slice(&outcome);
    }
    writer
        .write_all(&(label.len() as u64).to_le_bytes())
        .unwrap();
    writer.write_all(label.as_bytes()).unwrap();
    writer
        .write_all(&(record.len() as u64).to_le_bytes())
        .unwrap();
    writer.write_all(&record).unwrap();
}
fn expected(reader: &mut impl Read) -> (bool, usize, u64) {
    let mut flag = [0];
    reader.read_exact(&mut flag).unwrap();
    (flag[0] == 1, number(reader) as usize, number(reader))
}
fn run<P: Program>(language: &str, entry: &str, generated: &P, args: &[String]) {
    let grammar = fs::read_to_string(format!("{}/{language}/src/grammar.pest", args[3])).unwrap();
    let resolved = vybe_parser::compile(&grammar).unwrap();
    let mut reader = BufReader::new(fs::File::open(&args[2]).unwrap());
    let options = MatchOptions {
        max_steps: args[4].parse().unwrap(),
        max_rule_depth: args[5].parse().unwrap(),
    };
    let mut writer = BufWriter::new(io::stdout().lock());
    for _ in 0..number(&mut reader) {
        let label = string(&mut reader);
        let source = string(&mut reader);
        let raw = expected(&mut reader);
        case(
            &resolved,
            generated,
            entry,
            &label,
            &source,
            raw,
            options,
            &mut writer,
        );
        if language == "python" || language == "fortran" {
            let processed = expected(&mut reader);
            let source = if language == "python" {
                python_helper::preprocess_indentation(&source)
            } else {
                fortran_helper::rewrite_hollerith(&source)
            };
            case(
                &resolved,
                generated,
                entry,
                &format!("preprocessed:{label}"),
                &source,
                processed,
                options,
                &mut writer,
            );
        }
    }
    writer.flush().unwrap();
}
fn main() {
    std::thread::Builder::new()
        .stack_size(64 * 1024 * 1024)
        .spawn(|| {
            let args = env::args().collect::<Vec<_>>();
            assert_eq!(args.len(), 6);
            macro_rules! language {
                ($name:ident, $entry:ident) => {
                    run(
                        stringify!($name),
                        stringify!($entry),
                        &vybe_parser_generated_tests::all::$name::Parser,
                        &args,
                    )
                };
            }
            match args[1].as_str() {
                "js" => language!(js, program),
                "php" => language!(php, program),
                "python" => language!(python, program),
                "c" => language!(c, program),
                "csharp" => language!(csharp, program),
                "java" => language!(java, program),
                "dart" => language!(dart, program),
                "go" => language!(go, program),
                "kotlin" => language!(kotlin, program),
                "lua" => language!(lua, chunk),
                "powershell" => language!(powershell, program),
                "ruby" => language!(ruby, program),
                "cobol" => language!(cobol, program),
                "fortran" => language!(fortran, program),
                "pascal" => language!(pascal, program),
                "vb" => language!(vb, program),
                "washm" => language!(washm, program),
                "wast" => language!(wast, program),
                _ => panic!("unknown language"),
            }
        })
        .unwrap()
        .join()
        .unwrap();
}
