use vybe_bytecode::Chunk;

fn compile_dart(src: &str) -> Vec<Chunk> {
    let module = vybe_compiler::languages::dart::parse(src).expect("dart parse");
    let profile =
        vybe_compiler::profile::parse_profile(vybe_compiler::languages::dart::profile_source())
            .expect("dart profile");
    vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("dart compile")
}

fn compile_ruby(src: &str) -> Vec<Chunk> {
    let module = vybe_compiler::languages::ruby::parse(src).expect("ruby parse");
    let profile =
        vybe_compiler::profile::parse_profile(vybe_compiler::languages::ruby::profile_source())
            .expect("ruby profile");
    vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("ruby compile")
}

fn compile_cobol(src: &str) -> Vec<Chunk> {
    let module = vybe_compiler::languages::cobol::parse(src).expect("cobol parse");
    let profile =
        vybe_compiler::profile::parse_profile(vybe_compiler::languages::cobol::profile_source())
            .expect("cobol profile");
    vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("cobol compile")
}

fn compile_fortran(src: &str) -> Vec<Chunk> {
    let module = vybe_compiler::languages::fortran::parse(src).expect("fortran parse");
    let profile =
        vybe_compiler::profile::parse_profile(vybe_compiler::languages::fortran::profile_source())
            .expect("fortran profile");
    vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("fortran compile")
}

fn compile_vb(src: &str) -> Vec<Chunk> {
    let module = vybe_compiler::languages::vb::parse(src).expect("vb parse");
    let profile =
        vybe_compiler::profile::parse_profile(vybe_compiler::languages::vb::profile_source())
            .expect("vb profile");
    vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("vb compile")
}

fn compile_pascal(src: &str) -> Vec<Chunk> {
    let module = vybe_compiler::languages::pascal::parse(src).expect("pascal parse");
    let profile =
        vybe_compiler::profile::parse_profile(vybe_compiler::languages::pascal::profile_source())
            .expect("pascal profile");
    vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("pascal compile")
}

fn chunk_named<'a>(chunks: &'a [Chunk], name: &str) -> &'a Chunk {
    chunks
        .iter()
        .find(|chunk| chunk.name.eq_ignore_ascii_case(name))
        .unwrap_or_else(|| {
            panic!(
                "missing chunk {name}; got {:?}",
                chunks.iter().map(|chunk| &chunk.name).collect::<Vec<_>>()
            )
        })
}

#[test]
fn shared_generator_emission_marks_dart_functions() {
    let chunks = compile_dart(
        r#"
Iterable<int> more() sync* {
  yield 2;
}

Iterable<int> count() sync* {
  yield 1;
  yield* more();
}
"#,
    );
    assert!(chunk_named(&chunks, "more").is_generator);
    assert!(chunk_named(&chunks, "count").is_generator);
}

#[test]
fn shared_generator_emission_marks_ruby_methods() {
    let chunks = compile_ruby("def count\n  yield 1\n  yield 2\nend\n");
    assert!(chunk_named(&chunks, "count").is_generator);
}

#[test]
fn shared_generator_emission_marks_cobol_paragraphs() {
    let chunks = compile_cobol(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    STOP RUN.
GENERATOR-PARA.
    YIELD.
"#,
    );
    assert!(chunk_named(&chunks, "GENERATOR-PARA").is_generator);
}

#[test]
fn shared_generator_emission_marks_fortran_functions() {
    let chunks = compile_fortran(
        r#"
program test
contains
    function count() result(res)
        yield 1
    end function count
end program test
"#,
    );
    assert!(chunk_named(&chunks, "count").is_generator);
}

#[test]
fn shared_generator_emission_marks_vb_functions() {
    let chunks = compile_vb(
        r#"
Module Program
    Function Count()
        Yield 1
        Yield 2
    End Function
End Module
"#,
    );
    assert!(chunk_named(&chunks, "Count").is_generator);
}

#[test]
fn shared_generator_emission_marks_pascal_functions() {
    let chunks = compile_pascal(
        r#"
program T;
function Count: Integer;
begin
  yield 1;
end;
begin
end.
"#,
    );
    assert!(chunk_named(&chunks, "Count").is_generator);
}
#[test]
fn debug_gen_imports() {
    let src = r#"
function* gen() {
    yield 1;
    yield 2;
    yield 3;
}
let g = gen();
console.log(g.next().value);
"#;
    let module = vybe_compiler::languages::js::parse(src).expect("JS parse failed");
    let profile =
        vybe_compiler::profile::parse_profile(vybe_compiler::languages::js::profile_source())
            .expect("Failed to parse JS profile");
    let chunks = vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("JS compile failed");
    // After normalize_import_table: chunks[0] has all imports, others cleared
    // Verify gen chunk import indices map correctly to chunks[0]
    let ci_bytes = vybe_bytecode::opcode::Op::CALL_IMPORT.encode();
    let gen_chunk = &chunks[1];
    let code = &gen_chunk.code;
    let mut first_3_calls: Vec<(usize, u16, u8)> = Vec::new();
    let mut pos = 0;
    while pos + 6 < code.len() && first_3_calls.len() < 6 {
        if code[pos] == ci_bytes[0] && code[pos+1] == ci_bytes[1]
           && code[pos+2] == ci_bytes[2] && code[pos+3] == ci_bytes[3] {
            let idx = ((code[pos+4] as u16) << 8) | code[pos+5] as u16;
            let argc = code[pos+6];
            first_3_calls.push((pos, idx, argc));
            pos += 7;
        } else {
            pos += 1;
        }
    }
    eprintln!("chunks[0] has {} imports", chunks[0].imports.len());
    for (offset, idx, argc) in &first_3_calls {
        let import_name = if (*idx as usize) < chunks[0].imports.len() {
            let imp = &chunks[0].imports[*idx as usize];
            format!("{}.{}", imp.module, imp.name)
        } else {
            format!("OUT OF RANGE ({})", idx)
        };
        eprintln!("  gen @{:04}: call_import idx={} argc={} => {}", offset, idx, argc, import_name);
    }
    // Check header bytes
    eprintln!("gen chunk first 8 bytes: {:?}", &code[..8.min(code.len())]);
    // Check what's at the failing offset (around 1642-1656)
    let mut all_calls: Vec<(usize, u16, u8)> = Vec::new();
    let mut pos2 = 0;
    while pos2 + 6 < code.len() {
        if code[pos2] == ci_bytes[0] && code[pos2+1] == ci_bytes[1]
           && code[pos2+2] == ci_bytes[2] && code[pos2+3] == ci_bytes[3] {
            let idx = ((code[pos2+4] as u16) << 8) | code[pos2+5] as u16;
            let argc = code[pos2+6];
            all_calls.push((pos2, idx, argc));
            pos2 += 7;
        } else {
            pos2 += 1;
        }
    }
    eprintln!("\n--- Calls near offset 1640-1660 ---");
    for (offset, idx, argc) in &all_calls {
        if *offset >= 1630 && *offset <= 1670 {
            let import_name = if (*idx as usize) < chunks[0].imports.len() {
                let imp = &chunks[0].imports[*idx as usize];
                format!("{}.{}", imp.module, imp.name)
            } else {
                format!("OUT OF RANGE ({})", idx)
            };
            eprintln!("  gen @{:04}: call_import idx={} argc={} => {}", offset, idx, argc, import_name);
        }
    }
}
