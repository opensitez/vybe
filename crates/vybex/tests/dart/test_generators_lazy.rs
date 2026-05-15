use super::helpers::run_prints;

#[test]
fn sync_star_returns_continuation() {
    let out = run_prints(r#"
Iterable<int> count() sync* {
  yield 1;
  yield 2;
}

void main() {
  print(count());
}
"#);
    assert_eq!(out, ["[continuation]"]);
}

#[test]
fn sync_star_body_stays_lazy() {
    let out = run_prints(r#"
Iterable<int> loud() sync* {
  print('bad');
  yield 1;
}

void main() {
  var _ = loud();
  print('ok');
}
"#);
    assert_eq!(out, ["ok"]);
}