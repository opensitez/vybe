use super::helpers::compile_ok;

// ── ACCEPT FROM ENVIRONMENT ───────────────────────────────────

#[test] fn accept_environment_basic() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-home PIC X(200).
       PROCEDURE DIVISION.
           ACCEPT ws-home FROM ENVIRONMENT "HOME"
           DISPLAY ws-home
           STOP RUN.
"#);
}

#[test] fn accept_environment_path() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-path PIC X(1024).
       PROCEDURE DIVISION.
           ACCEPT ws-path FROM ENVIRONMENT "PATH"
           DISPLAY ws-path
           STOP RUN.
"#);
}

#[test] fn accept_environment_user() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-user PIC X(64).
       PROCEDURE DIVISION.
           ACCEPT ws-user FROM ENVIRONMENT "USER"
           DISPLAY "User: " ws-user
           STOP RUN.
"#);
}

#[test] fn accept_environment_name_keyword() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-env-name PIC X(30) VALUE "HOME".
       01 ws-env-val  PIC X(200).
       PROCEDURE DIVISION.
           ACCEPT ws-env-val FROM ENVIRONMENT NAME ws-env-name
           DISPLAY ws-env-val
           STOP RUN.
"#);
}

#[test] fn accept_environment_name_variable() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-var-name PIC X(20) VALUE "TMPDIR".
       01 ws-var-val  PIC X(200).
       PROCEDURE DIVISION.
           ACCEPT ws-var-val FROM ENVIRONMENT NAME ws-var-name
           IF ws-var-val = SPACES
               DISPLAY "not set"
           ELSE
               DISPLAY ws-var-val
           END-IF
           STOP RUN.
"#);
}

#[test] fn accept_environment_custom_var() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-db-host PIC X(100).
       01 ws-db-port PIC X(10).
       PROCEDURE DIVISION.
           ACCEPT ws-db-host FROM ENVIRONMENT "DB_HOST"
           ACCEPT ws-db-port FROM ENVIRONMENT "DB_PORT"
           IF ws-db-host = SPACES
               MOVE "localhost" TO ws-db-host
           END-IF
           IF ws-db-port = SPACES
               MOVE "5432" TO ws-db-port
           END-IF
           DISPLAY ws-db-host
           DISPLAY ws-db-port
           STOP RUN.
"#);
}

#[test] fn accept_environment_missing_var() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-val PIC X(50).
       01 ws-status PIC X(10).
       PROCEDURE DIVISION.
           MOVE SPACES TO ws-val
           ACCEPT ws-val FROM ENVIRONMENT "NONEXISTENT_VAR_XYZ"
           IF ws-val = SPACES
               MOVE "missing" TO ws-status
           ELSE
               MOVE "found" TO ws-status
           END-IF
           DISPLAY ws-status
           STOP RUN.
"#);
}

#[test] fn accept_environment_multiple() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-home  PIC X(200).
       01 ws-user  PIC X(64).
       01 ws-shell PIC X(100).
       PROCEDURE DIVISION.
           ACCEPT ws-home  FROM ENVIRONMENT "HOME"
           ACCEPT ws-user  FROM ENVIRONMENT "USER"
           ACCEPT ws-shell FROM ENVIRONMENT "SHELL"
           DISPLAY "home:  " ws-home
           DISPLAY "user:  " ws-user
           DISPLAY "shell: " ws-shell
           STOP RUN.
"#);
}

#[test] fn accept_environment_and_display() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-app-mode PIC X(10).
       01 ws-log-level PIC X(10).
       PROCEDURE DIVISION.
           ACCEPT ws-app-mode  FROM ENVIRONMENT "APP_MODE"
           ACCEPT ws-log-level FROM ENVIRONMENT "LOG_LEVEL"
           EVALUATE ws-app-mode
               WHEN "production"
                   DISPLAY "Running in production"
               WHEN "staging"
                   DISPLAY "Running in staging"
               WHEN OTHER
                   DISPLAY "Running in dev mode"
           END-EVALUATE
           STOP RUN.
"#);
}

#[test] fn accept_environment_in_subroutine() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. main-prog.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-config PIC X(200).
       PROCEDURE DIVISION.
           CALL "get-config" USING ws-config
           DISPLAY ws-config
           STOP RUN.

       IDENTIFICATION DIVISION.
       PROGRAM-ID. get-config.
       DATA DIVISION.
       LINKAGE SECTION.
       01 lk-config PIC X(200).
       WORKING-STORAGE SECTION.
       01 ws-val PIC X(200).
       PROCEDURE DIVISION USING lk-config.
           ACCEPT ws-val FROM ENVIRONMENT "APP_CONFIG"
           IF ws-val = SPACES
               MOVE "default.cfg" TO lk-config
           ELSE
               MOVE ws-val TO lk-config
           END-IF
           GOBACK.
"#);
}

#[test] fn accept_environment_with_inspect() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-path     PIC X(500).
       01 ws-colon-ct PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           ACCEPT ws-path FROM ENVIRONMENT "PATH"
           INSPECT ws-path
               TALLYING ws-colon-ct FOR ALL ":"
           DISPLAY ws-colon-ct
           STOP RUN.
"#);
}

#[test] fn accept_environment_numeric() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-timeout-str PIC X(10).
       01 ws-timeout-val PIC 9(5) VALUE 30.
       PROCEDURE DIVISION.
           ACCEPT ws-timeout-str FROM ENVIRONMENT "TIMEOUT_SECS"
           IF ws-timeout-str NOT = SPACES
               MOVE FUNCTION NUMVAL(ws-timeout-str) TO ws-timeout-val
           END-IF
           DISPLAY ws-timeout-val
           STOP RUN.
"#);
}

// ── ACCEPT FROM DATE / TIME / DAY ────────────────────────────

#[test] fn accept_from_date() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-date PIC 9(6).
       PROCEDURE DIVISION.
           ACCEPT ws-date FROM DATE
           DISPLAY ws-date
           STOP RUN.
"#);
}

#[test] fn accept_from_date_yyyymmdd() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-date PIC 9(8).
       PROCEDURE DIVISION.
           ACCEPT ws-date FROM DATE YYYYMMDD
           DISPLAY ws-date
           STOP RUN.
"#);
}

#[test] fn accept_from_time() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-time PIC 9(8).
       PROCEDURE DIVISION.
           ACCEPT ws-time FROM TIME
           DISPLAY ws-time
           STOP RUN.
"#);
}

#[test] fn accept_from_day() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-day PIC 9(5).
       PROCEDURE DIVISION.
           ACCEPT ws-day FROM DAY
           DISPLAY ws-day
           STOP RUN.
"#);
}

#[test] fn accept_from_day_yyyyddd() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-day PIC 9(7).
       PROCEDURE DIVISION.
           ACCEPT ws-day FROM DAY YYYYDDD
           DISPLAY ws-day
           STOP RUN.
"#);
}

#[test] fn accept_from_day_of_week() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-dow PIC 9.
       PROCEDURE DIVISION.
           ACCEPT ws-dow FROM DAY-OF-WEEK
           EVALUATE ws-dow
               WHEN 1 DISPLAY "Monday"
               WHEN 2 DISPLAY "Tuesday"
               WHEN 3 DISPLAY "Wednesday"
               WHEN 4 DISPLAY "Thursday"
               WHEN 5 DISPLAY "Friday"
               WHEN 6 DISPLAY "Saturday"
               WHEN 7 DISPLAY "Sunday"
           END-EVALUATE
           STOP RUN.
"#);
}

#[test] fn accept_date_and_time_combined() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-date PIC 9(6).
       01 ws-time PIC 9(8).
       01 ws-day  PIC 9(5).
       PROCEDURE DIVISION.
           ACCEPT ws-date FROM DATE
           ACCEPT ws-time FROM TIME
           ACCEPT ws-day  FROM DAY
           DISPLAY ws-date
           DISPLAY ws-time
           DISPLAY ws-day
           STOP RUN.
"#);
}
