use super::helpers::compile_ok;

// Basic try/except

#[test]
fn simple_try_except() {
    compile_ok("try:\n    x = 1\nexcept:\n    pass\n");
}

#[test]
fn try_except_with_name() {
    compile_ok("try:\n    x = 1 / 0\nexcept Exception as e:\n    print(e)\n");
}

// Typed exception handlers

#[test]
fn typed_except_single() {
    compile_ok(
        r#"
try:
    x = int("abc")
except ValueError:
    print("bad value")
"#,
    );
}

#[test]
fn typed_except_multiple() {
    compile_ok(
        r#"
try:
    x = 1 / 0
except ValueError:
    print("value error")
except TypeError:
    print("type error")
except ZeroDivisionError:
    print("division by zero")
"#,
    );
}

#[test]
fn typed_except_with_names() {
    compile_ok(
        r#"
try:
    result = dangerous_operation()
except ValueError as ve:
    print("ValueError:", ve)
except TypeError as te:
    print("TypeError:", te)
except Exception as e:
    print("Other:", e)
"#,
    );
}

#[test]
fn typed_except_with_catch_all() {
    compile_ok(
        r#"
try:
    x = 1
except ValueError:
    print("value error")
except:
    print("catch all")
"#,
    );
}

// Try/except/else/finally

#[test]
fn try_except_else() {
    compile_ok(
        r#"
try:
    x = 1
except:
    print("error")
else:
    print("no error")
"#,
    );
}

#[test]
fn try_except_finally() {
    compile_ok(
        r#"
try:
    f = open("file.txt")
except:
    print("error")
finally:
    print("cleanup")
"#,
    );
}

#[test]
fn try_except_else_finally() {
    compile_ok(
        r#"
try:
    x = 1
except ValueError:
    print("value error")
else:
    print("success")
finally:
    print("done")
"#,
    );
}

// Raise

#[test]
fn raise_typed() {
    compile_ok("raise ValueError()\n");
}

#[test]
fn raise_with_message() {
    compile_ok(r#"raise ValueError("bad input")"#);
}

#[test]
fn raise_from() {
    compile_ok(
        r#"
try:
    x = 1 / 0
except ZeroDivisionError as e:
    raise ValueError("invalid") from e
"#,
    );
}

// Nested try/except

#[test]
fn nested_try() {
    compile_ok(
        r#"
try:
    try:
        x = 1 / 0
    except ZeroDivisionError:
        print("inner")
        raise ValueError("converted")
except ValueError:
    print("outer")
"#,
    );
}
