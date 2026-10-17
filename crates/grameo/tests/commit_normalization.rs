use grameo::{builder::Builder, engine, MatchOptions, Span};

#[derive(Default)]
struct Normalizer { values: Vec<String>, commits: usize }
impl Builder for Normalizer {
    fn checkpoint(&self) -> usize { self.values.len() }
    fn rollback(&mut self, mark: usize) { self.values.truncate(mark); }
    fn begin(&mut self, _: usize, _: usize) {}
    fn finish(&mut self, _: usize, span: Span, source: &str) {
        self.values.push(source[span.start..span.end].to_owned());
    }
    fn committed(&mut self, _: &str, _: usize) {
        self.commits += 1;
        for value in &mut self.values { value.make_ascii_uppercase(); }
    }
}

#[test]
fn normalization_runs_once_after_success_and_never_for_failed_speculation() {
    let grammar = grameo::compile("root = { (word ~ \"!\" | word) ~ EOI } word = @{ ASCII_ALPHA+ }").unwrap();
    for captures in [false, true] {
        let mut builder = Normalizer::default();
        if captures {
            engine::build_capture_program(&grammar, "root", "hello", MatchOptions::default(), &mut builder).unwrap();
        } else {
            engine::build_ast_program(&grammar, "root", "hello", MatchOptions::default(), &mut builder).unwrap();
        }
        assert_eq!(builder.commits, 1);
        assert_eq!(builder.values, ["HELLO", "HELLO"]);
        let before = builder.values.clone();
        let failed = if captures {
            engine::build_capture_program(&grammar, "root", "hello?", MatchOptions::default(), &mut builder).map(|_| ())
        } else {
            engine::build_ast_program(&grammar, "root", "hello?", MatchOptions::default(), &mut builder).map(|_| ())
        };
        assert!(failed.is_err());
        assert_eq!(builder.commits, 1);
        assert_eq!(builder.values, before);
    }
}
