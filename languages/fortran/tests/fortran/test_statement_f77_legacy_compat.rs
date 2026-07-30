use super::helpers::run_prints;

#[test]
fn statement_f77_legacy_compat_explicit_type_suffix() {
    let out = run_prints(
        "program statement_f77_legacy_compat_explicit_type_suffix\n\
            integer*4 i\n\
            i = 1\n\
            print *, i\n\
         end program statement_f77_legacy_compat_explicit_type_suffix\n",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn statement_f77_legacy_compat_fixed_width_integers() {
    let out = run_prints(
        "program statement_f77_legacy_compat_fixed_width_integers\n\
            integer*2 short\n\
            integer*8 long\n\
            short = 2\n\
            long = 3\n\
            print *, short\n\
            print *, long\n\
         end program statement_f77_legacy_compat_fixed_width_integers\n",
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn statement_f77_legacy_compat_entry_statement() {
    let out = run_prints(
        "program statement_f77_legacy_compat_entry_statement\n\
            integer :: sum\n\
            sum = value(1, 2) + value2(3, 4)\n\
            print *, sum\n\
        contains\n\
            integer function value(a, b)\n\
                integer a, b\n\
                value = a + b\n\
                return\n\
            entry value2(a, b)\n\
                value2 = a * b\n\
            end function value\n\
        end program statement_f77_legacy_compat_entry_statement\n",
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn statement_f77_legacy_compat_assigned_goto() {
    let out = run_prints(
        "program statement_f77_legacy_compat_assigned_goto\n\
            integer label\n\
            assign 20 to label\n\
            goto label\n\
            print *, 'skip'\n\
20          print *, 'hit'\n\
        end program statement_f77_legacy_compat_assigned_goto\n",
    );
    assert_eq!(out, vec!["hit"]);
}

#[test]
fn statement_f77_legacy_compat_arithmetic_if() {
    let out = run_prints(
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
    assert_eq!(out, vec!["neg"]);
}

#[test]
fn statement_f77_legacy_compat_data_statement_init() {
    let out = run_prints(
        "program statement_f77_legacy_compat_data_statement_init\n\
            integer i, j\n\
            data i /1/, j /2/\n\
            print *, i + j\n\
        end program statement_f77_legacy_compat_data_statement_init\n",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn statement_f77_legacy_compat_hollerith_like_literal() {
    let out = run_prints(
        "program statement_f77_legacy_compat_hollerith_like_literal\n\
            character*12 c\n\
            data c / 'HELLOWORLD  ' /\n\
            print *, trim(c)\n\
        end program statement_f77_legacy_compat_hollerith_like_literal\n",
    );
    assert_eq!(out, vec!["HELLOWORLD"]);
}

#[test]
fn statement_f77_legacy_compat_logical_and_goto_label_references() {
    let out = run_prints(
        "program statement_f77_legacy_compat_logical_and_goto_label_references\n\
            integer i\n\
            i = 1\n\
            if (i .gt. 0) goto 100\n\
            print *, 'bad'\n\
100         print *, 'good'\n\
        end program statement_f77_legacy_compat_logical_and_goto_label_references\n",
    );
    assert_eq!(out, vec!["good"]);
}

#[test]
fn statement_f77_legacy_compat_computed_goto() {
    let out = run_prints(
        "program statement_f77_legacy_compat_computed_goto\n\
            integer n\n\
            n = 2\n\
            go to (10, 20, 30), n\n\
10          print *, 10\n\
            go to 99\n\
20          print *, 20\n\
            go to 99\n\
30          print *, 30\n+99          continue\n\
        end program statement_f77_legacy_compat_computed_goto\n",
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn statement_f77_legacy_compat_labeled_do() {
    let out = run_prints(
        "program statement_f77_legacy_compat_labeled_do\n\
            integer i\n\
            integer sum\n\
            sum = 0\n\
            do 10 i = 1, 3\n\
                sum = sum + i\n\
10          continue\n\
            print *, sum\n\
        end program statement_f77_legacy_compat_labeled_do\n",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn statement_f77_legacy_compat_common_blocks() {
    let out = run_prints(
        "program statement_f77_legacy_compat_common_blocks\n\
            integer a, b\n\
            common /legacy1/ a, b\n\
            integer sum\n\
            common /legacy2/ sum\n\
            a = 1\n\
            b = 2\n\
            sum = a + b\n\
            print *, sum\n\
        end program statement_f77_legacy_compat_common_blocks\n",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn statement_f77_legacy_compat_data_repeat_syntax() {
    let out = run_prints(
        "program statement_f77_legacy_compat_data_repeat_syntax\n\
            integer i, j(3)\n\
            data i /3*2/\n\
            data (j(k), k=1,3) /1,2,3/\n\
            print *, i\n\
            print *, j(1)\n\
            print *, j(2)\n\
            print *, j(3)\n\
        end program statement_f77_legacy_compat_data_repeat_syntax\n",
    );
    assert_eq!(out, vec!["6", "1", "2", "3"]);
}

#[test]
fn statement_f77_legacy_compat_assigned_goto_multiple_targets() {
    let out = run_prints(
        "program statement_f77_legacy_compat_assigned_goto_multiple_targets\n\
            integer label\n\
            assign 10 to label\n\
            integer i\n\
            i = 1\n\
            go to label\n\
            print *, 'other'\n\
10          print *, 'matched'\n\
        end program statement_f77_legacy_compat_assigned_goto_multiple_targets\n",
    );
    assert_eq!(out, vec!["matched"]);
}

#[test]
fn statement_f77_legacy_compat_hollerith_literal() {
    let out = run_prints(
        "program statement_f77_legacy_compat_hollerith_literal\n\
            character*4 c\n\
            c = 4hABCD\n\
            print *, c\n\
        end program statement_f77_legacy_compat_hollerith_literal\n",
    );
    assert_eq!(out, vec!["ABCD"]);
}

#[test]
fn statement_f77_legacy_compat_statement_function() {
    let out = run_prints(
        "program statement_f77_legacy_compat_statement_function\n\
            integer n\n\
            integer square\n\
            n = 5\n\
            square(x) = x * x\n\
            print *, square(n)\n\
            print *, square(3)\n\
        end\n",
    );
    assert_eq!(out, vec!["25", "9"]);
}

#[test]
fn statement_f77_legacy_compat_format_write_variants() {
    let out = run_prints(
        "program statement_f77_legacy_compat_format_write_variants\n\
            integer i\n\
            i = 4\n\
            print 10, i\n\
10          format (I1)\n\
        end program statement_f77_legacy_compat_format_write_variants\n",
    );
    assert_eq!(out[0].trim(), "4");
}
