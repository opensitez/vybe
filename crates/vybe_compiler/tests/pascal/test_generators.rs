use super::helpers::run_pascal;

#[test]
fn function_yield_returns_continuation() {
    let out = run_pascal(r#"
program T;

function Count: Integer;
begin
  yield 1;
  yield 2;
end;

begin
  WriteLn(Count());
end.
"#);

    assert_eq!(out, ["[continuation]"]);
}

#[test]
fn function_yield_body_stays_lazy() {
    let out = run_pascal(r#"
program T;

function Loud: Integer;
begin
  WriteLn('bad');
  yield 1;
end;

begin
  Loud();
  WriteLn('ok');
end.
"#);

    assert_eq!(out, ["ok"]);
}