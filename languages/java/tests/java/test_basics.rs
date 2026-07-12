use crate::helpers::run_main;

#[test]
fn hello_world_println() {
    let out = run_main(r#"System.out.println("hello world");"#);
    assert_eq!(out, vec!["hello world"]);
}

#[test]
fn integer_variable_declared_and_printed() {
    let out = run_main("int x = 42; System.out.println(x);");
    assert_eq!(out, vec!["42"]);
}

#[test]
fn final_local_cannot_be_reassigned() {
    let out = run_main("final int x = 7; System.out.println(x);");
    assert_eq!(out, vec!["7"]);
}

#[test]
fn multiple_variables_declared_on_one_line() {
    let out = run_main("int a = 1, b = 2, c = 3; System.out.println(a + b + c);");
    assert_eq!(out, vec!["6"]);
}

#[test]
fn arithmetic_add_sub_mul_div() {
    let out = run_main(
        "System.out.println(3 + 4); System.out.println(10 - 3); System.out.println(6 * 7); System.out.println(20 / 4);",
    );
    assert_eq!(out, vec!["7", "7", "42", "5"]);
}

#[test]
fn boolean_literals_true_and_false() {
    let out = run_main(
        "boolean t = true; boolean f = false; System.out.println(t); System.out.println(f);",
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn prefix_increment_updates_before_use() {
    let out = run_main("int x = 5; System.out.println(++x); System.out.println(x);");
    assert_eq!(out, vec!["6", "6"]);
}

#[test]
fn postfix_increment_updates_after_use() {
    let out = run_main("int x = 5; System.out.println(x++); System.out.println(x);");
    assert_eq!(out, vec!["5", "6"]);
}

#[test]
fn compound_addition_assignment() {
    let out = run_main("int x = 10; x += 5; System.out.println(x);");
    assert_eq!(out, vec!["15"]);
}
