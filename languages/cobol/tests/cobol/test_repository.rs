use super::helpers::compile_ok;

// ── REPOSITORY paragraph — CLASS ─────────────────────────────

#[test]
fn repository_class_basic() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           CLASS MyClass AS "MyClass".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-obj OBJECT REFERENCE MyClass.
       PROCEDURE DIVISION.
           DISPLAY "repository class declared"
           STOP RUN.
"#,
    );
}

#[test]
fn repository_multiple_classes() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           CLASS Animal   AS "Animal"
           CLASS Dog      AS "Dog"
           CLASS Cat      AS "Cat".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-animal OBJECT REFERENCE Animal.
       01 ws-dog    OBJECT REFERENCE Dog.
       PROCEDURE DIVISION.
           DISPLAY "multi-class repository"
           STOP RUN.
"#,
    );
}

#[test]
fn repository_class_and_interface() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           CLASS     Shape     AS "Shape"
           INTERFACE Drawable  AS "Drawable".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-shape OBJECT REFERENCE Shape.
       PROCEDURE DIVISION.
           DISPLAY "class and interface in repository"
           STOP RUN.
"#,
    );
}

// ── REPOSITORY — FUNCTION ALL INTRINSIC ──────────────────────

#[test]
fn repository_function_all_intrinsic() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           FUNCTION ALL INTRINSIC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9(5)V99.
       01 ws-text   PIC X(20) VALUE "hello world".
       PROCEDURE DIVISION.
           COMPUTE ws-result = SQRT(16)
           DISPLAY ws-result
           DISPLAY UPPER-CASE(ws-text)
           STOP RUN.
"#,
    );
}

#[test]
fn repository_function_all_intrinsic_math() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           FUNCTION ALL INTRINSIC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-pi      PIC 9V9(8).
       01 ws-e       PIC 9V9(8).
       01 ws-abs-val PIC 99V99.
       PROCEDURE DIVISION.
           COMPUTE ws-pi      = ACOS(-1)
           COMPUTE ws-e       = EXP(1)
           COMPUTE ws-abs-val = ABS(-3.14)
           DISPLAY ws-pi
           DISPLAY ws-e
           DISPLAY ws-abs-val
           STOP RUN.
"#,
    );
}

#[test]
fn repository_function_all_intrinsic_string() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           FUNCTION ALL INTRINSIC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-src  PIC X(20) VALUE "  hello world  ".
       01 ws-trim PIC X(20).
       01 ws-up   PIC X(20).
       01 ws-rev  PIC X(20).
       PROCEDURE DIVISION.
           MOVE TRIM(ws-src)         TO ws-trim
           MOVE UPPER-CASE(ws-src)   TO ws-up
           MOVE REVERSE(ws-src)      TO ws-rev
           DISPLAY ws-trim
           DISPLAY ws-up
           DISPLAY ws-rev
           STOP RUN.
"#,
    );
}

// ── REPOSITORY — specific FUNCTION entries ────────────────────

#[test]
fn repository_specific_function() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           FUNCTION SQRT
           FUNCTION ABS
           FUNCTION MOD.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-val PIC 9(5)V99.
       PROCEDURE DIVISION.
           COMPUTE ws-val = SQRT(25)
           DISPLAY ws-val
           COMPUTE ws-val = ABS(-7)
           DISPLAY ws-val
           STOP RUN.
"#,
    );
}

#[test]
fn repository_function_with_alias() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           FUNCTION UPPER-CASE AS "UPPER-CASE"
           FUNCTION LOWER-CASE AS "LOWER-CASE"
           FUNCTION TRIM       AS "TRIM".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-text PIC X(20) VALUE "Hello World".
       01 ws-up   PIC X(20).
       01 ws-lo   PIC X(20).
       PROCEDURE DIVISION.
           MOVE UPPER-CASE(ws-text) TO ws-up
           MOVE LOWER-CASE(ws-text) TO ws-lo
           DISPLAY ws-up
           DISPLAY ws-lo
           STOP RUN.
"#,
    );
}

// ── REPOSITORY in CLASS-ID definitions ───────────────────────

#[test]
fn repository_in_class() {
    compile_ok(
        r#"
       CLASS-ID. Calculator.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           FUNCTION ALL INTRINSIC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-precision PIC 99 VALUE 8.
       METHOD-ID. square-root.
       DATA DIVISION.
       LOCAL-STORAGE SECTION.
       01 ls-result PIC 9(5)V9(8).
       LINKAGE SECTION.
       01 lk-input  PIC 9(5)V9(8).
       01 lk-result PIC 9(5)V9(8).
       PROCEDURE DIVISION USING lk-input RETURNING lk-result.
           COMPUTE lk-result = SQRT(lk-input)
           GOBACK.
       END METHOD square-root.
       END CLASS Calculator.
"#,
    );
}

#[test]
fn repository_class_hierarchy() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           CLASS Vehicle  AS "Vehicle"
           CLASS Car      AS "Car"
           CLASS Truck    AS "Truck"
           CLASS Fleet    AS "Fleet".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-vehicle OBJECT REFERENCE Vehicle.
       01 ws-car     OBJECT REFERENCE Car.
       01 ws-truck   OBJECT REFERENCE Truck.
       PROCEDURE DIVISION.
           DISPLAY "fleet management system"
           STOP RUN.
"#,
    );
}

// ── REPOSITORY with both FUNCTION and CLASS ───────────────────

#[test]
fn repository_mixed() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           FUNCTION ALL INTRINSIC
           CLASS Connection AS "Connection"
           CLASS ResultSet  AS "ResultSet"
           INTERFACE Closeable AS "Closeable".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-conn OBJECT REFERENCE Connection.
       01 ws-rs   OBJECT REFERENCE ResultSet.
       01 ws-len  PIC 99.
       PROCEDURE DIVISION.
           COMPUTE ws-len = LENGTH("hello")
           DISPLAY ws-len
           STOP RUN.
"#,
    );
}

#[test]
fn repository_in_module_with_invoke() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           CLASS StringUtil AS "StringUtil"
           FUNCTION ALL INTRINSIC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-util  OBJECT REFERENCE StringUtil.
       01 ws-input PIC X(30) VALUE "hello world".
       01 ws-len   PIC 99.
       PROCEDURE DIVISION.
           COMPUTE ws-len = LENGTH(TRIM(ws-input))
           DISPLAY ws-len
           STOP RUN.
"#,
    );
}
