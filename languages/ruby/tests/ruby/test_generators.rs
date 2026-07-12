use super::helpers::run_ruby;

#[test]
fn yield_method_returns_continuation() {
    let out = run_ruby("def count\n  yield 1\n  yield 2\nend\nputs count()\n");
    assert_eq!(out, vec!["[continuation]"]);
}

#[test]
fn yield_method_body_stays_lazy() {
    let out = run_ruby("def loud\n  puts 'bad'\n  yield 1\nend\ng = loud\nputs 'ok'\n");
    assert_eq!(out, vec!["ok"]);
}
