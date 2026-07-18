use super::helpers::compile_ok;

#[test]
fn statement_f77_legacy_compat_explicit_type_suffix() {
    compile_ok(
        "program statement_f77_legacy_compat_explicit_type_suffix\n\
            integer*4 i\n\
            i = 1\n\
            print *, i\n\
         end program statement_f77_legacy_compat_explicit_type_suffix\n",
    );
}

#[test]
fn statement_f77_legacy_compat_fixed_width_integers() {
    compile_ok(
        "program statement_f77_legacy_compat_fixed_width_integers\n\
            integer*2 short\n\
            integer*8 long\n\
            short = 2\n\
            long = 3\n\
            print *, short\n\
            print *, long\n\
         end program statement_f77_legacy_compat_fixed_width_integers\n",
    );
}

#[test]
fn statement_f77_legacy_compat_entry_statement() {
    compile_ok(
        "program statement_f77_legacy_compat_entry_statement\n\
            integer :: sum\n\
            sum = 0\n\
            print *, value(1, 2)\n\
            end\n\
        contains\n            integer function value(a, b)\n                integer a, b\n                value = a + b\n                return\n            entry value2(a, b)\n                value2 = a * b\n            end function value\n        end program statement_f77_legacy_compat_entry_statement\n",
    );
}

#[test]
fn statement_f77_legacy_compat_assigned_goto() {
    compile_ok(
        "program statement_f77_legacy_compat_assigned_goto\n\
            integer label\n\
            assign 20 to label\n\
            goto label\n\
            print *, 'skip'\n\
20          print *, 'hit'\n\
        end program statement_f77_legacy_compat_assigned_goto\n",
    );
}

#[test]
fn statement_f77_legacy_compat_arithmetic_if() {
    compile_ok(
        "program statement_f77_legacy_compat_arithmetic_if\n\
            real x\n\
            x = -1.0\n\
            if (x) 10, 20, 30\n\
10          print *, 'neg'\n\
            stop 0\n\
20          print *, 'zero'\n\
            stop 0\n\
30          print *, 'pos'\n\
        end program statement_f77_legacy_compat_arithmetic_if\n",
    );
}

#[test]
fn statement_f77_legacy_compat_data_statement_init() {
    compile_ok(
        "program statement_f77_legacy_compat_data_statement_init\n\
            integer i, j\n\
            data i /1/, j /2/\n\
            print *, i + j\n\
        end program statement_f77_legacy_compat_data_statement_init\n",
    );
}

#[test]
fn statement_f77_legacy_compat_hollerith_like_literal() {
    compile_ok(
        "program statement_f77_legacy_compat_hollerith_like_literal\n\
            character*12 c\n\
            data c / 'HELLOWORLD  ' /\n\
            print *, c\n\
        end program statement_f77_legacy_compat_hollerith_like_literal\n",
    );
}

#[test]
fn statement_f77_legacy_compat_logical_and_goto_label_references() {
    compile_ok(
        "program statement_f77_legacy_compat_logical_and_goto_label_references\n\
            integer i\n\
            i = 1\n\
            if (i .gt. 0) goto 100\n\
            print *, 'bad'\n\
100         print *, 'good'\n\
        end program statement_f77_legacy_compat_logical_and_goto_label_references\n",
    );
}

#[test]
fn statement_f77_legacy_compat_computed_goto() {
    compile_ok(
        "program statement_f77_legacy_compat_computed_goto\n\
            integer n\n\
            n = 2\n\
            go to (10, 20, 30), n\n\
10          print *, 10\n\
            go to 99\n\
20          print *, 20\n\
            go to 99\n\
30          print *, 30\n99          continue\n\
        end program statement_f77_legacy_compat_computed_goto\n",
    );
}
