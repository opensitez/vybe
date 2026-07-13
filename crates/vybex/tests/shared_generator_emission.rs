use vybe_bytecode::Chunk;

fn compile_dart(src: &str) -> Vec<Chunk> {
    let module = vybe_language_dart::parse(src).expect("dart parse");
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_dart::profile_source())
            .expect("dart profile");
    vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("dart compile")
}

fn compile_ruby(src: &str) -> Vec<Chunk> {
    let module = vybe_language_ruby::parse(src).expect("ruby parse");
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_ruby::profile_source())
            .expect("ruby profile");
    vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("ruby compile")
}

fn compile_cobol(src: &str) -> Vec<Chunk> {
    let module = vybe_language_cobol::parse(src).expect("cobol parse");
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_cobol::profile_source())
            .expect("cobol profile");
    vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("cobol compile")
}

fn compile_fortran(src: &str) -> Vec<Chunk> {
    let module = vybe_language_fortran::parse(src).expect("fortran parse");
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_fortran::profile_source())
            .expect("fortran profile");
    vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("fortran compile")
}

fn compile_vb(src: &str) -> Vec<Chunk> {
    let module = vybe_language_vb::parse(src).expect("vb parse");
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_vb::profile_source())
            .expect("vb profile");
    vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("vb compile")
}

fn compile_pascal(src: &str) -> Vec<Chunk> {
    let module = vybe_language_pascal::parse(src).expect("pascal parse");
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_pascal::profile_source())
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
