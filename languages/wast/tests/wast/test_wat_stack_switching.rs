//! Stack switching (typed continuations) proposal — cont types, cont.new,
//! suspend, resume, cont.bind. Checks the instruction/type syntax parses.
use super::helpers::parse_ok;

#[test]
fn cont_type_definition() {
    parse_ok(r#"(module (type $ft (func)) (type $ct (cont $ft)))"#);
}
#[test]
fn tag_for_suspend() {
    parse_ok(r#"(module (tag $yield (param i32)))"#);
}
#[test]
fn cont_new_instruction() {
    parse_ok(
        r#"(module (type $ft (func)) (type $ct (cont $ft))
          (func $f)
          (func (export "_start") ref.func $f cont.new $ct drop))"#,
    );
}
#[test]
fn suspend_instruction() {
    parse_ok(
        r#"(module (tag $yield (param i32))
          (func (export "_start") i32.const 5 suspend $yield))"#,
    );
}
#[test]
fn resume_instruction() {
    parse_ok(
        r#"(module (type $ft (func)) (type $ct (cont $ft)) (tag $yield)
          (func $f)
          (func (export "_start") ref.func $f cont.new $ct resume $ct (on $yield $h))
          (func $h))"#,
    );
}
#[test]
fn cont_bind_instruction() {
    parse_ok(
        r#"(module (type $ft (func (param i32))) (type $ct (cont $ft))
          (type $ft2 (func)) (type $ct2 (cont $ft2))
          (func $f (param i32))
          (func (export "_start") ref.func $f cont.new $ct
            i32.const 1 cont.bind $ct $ct2 drop))"#,
    );
}
