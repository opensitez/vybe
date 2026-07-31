use crate::helpers;

#[test]
fn test_oo_invoke_basic() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. OO-INVOKE.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           CLASS ACCOUNT IS "Account".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 OBJ-REF USAGE OBJECT REFERENCE ACCOUNT.
       PROCEDURE DIVISION.
           INVOKE OBJ-REF "methodName".
           DISPLAY "INVOKE PARSED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_oo_invoke_factory() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. OO-FACTORY.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           CLASS ACCOUNT IS "Account".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 FACTORY-REF USAGE OBJECT REFERENCE FACTORY OF ACCOUNT.
       01 OBJ-REF USAGE OBJECT REFERENCE ACCOUNT.
       PROCEDURE DIVISION.
           INVOKE ACCOUNT "FACTORY" RETURNING FACTORY-REF.
           INVOKE FACTORY-REF "NEW" RETURNING OBJ-REF.
           DISPLAY "FACTORY PARSED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_oo_invoke_returning() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. OO-RET.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           CLASS ACCOUNT IS "Account".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 OBJ-REF USAGE OBJECT REFERENCE ACCOUNT.
       01 RES PIC 9(4).
       PROCEDURE DIVISION.
           INVOKE OBJ-REF "getBalance" RETURNING RES.
           DISPLAY "RETURNING PARSED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_oo_invoke_using() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. OO-USING.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           CLASS ACCOUNT IS "Account".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 OBJ-REF USAGE OBJECT REFERENCE ACCOUNT.
       01 AMT PIC 9(4) VALUE 100.
       PROCEDURE DIVISION.
           INVOKE OBJ-REF "deposit" USING BY VALUE AMT.
           DISPLAY "USING PARSED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_oo_invoke_multiple_arguments() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. OO-ARGS.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           CLASS ACCOUNT IS "Account".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 OBJ-REF USAGE OBJECT REFERENCE ACCOUNT.
       01 AMT1 PIC 9(4) VALUE 100.
       01 AMT2 PIC 9(4) VALUE 20.
       PROCEDURE DIVISION.
           INVOKE OBJ-REF "transfer" USING BY VALUE AMT1 BY VALUE AMT2.
           DISPLAY "OO-ARGS PARSED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_oo_invoke_returning_with_arguments() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. OO-RETARGS.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           CLASS ACCOUNT IS "Account".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 OBJ-REF USAGE OBJECT REFERENCE ACCOUNT.
       01 AMOUNT PIC 9(4) VALUE 200.
       01 RES PIC 9(4) VALUE ZERO.
       PROCEDURE DIVISION.
           INVOKE OBJ-REF "setBalance" USING BY VALUE AMOUNT RETURNING RES.
           DISPLAY "OO-RETARGS PARSED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}
