use super::helpers::compile_ok;

// ── Range selectors ───────────────────────────────────────────

#[test]
fn case_range_low() {
    compile_ok(
        r#"
program test
    integer :: n = 3
    select case (n)
    case (1:5)
        print *, 'low'
    case (6:10)
        print *, 'high'
    end select
end program test
"#,
    );
}

#[test]
fn case_range_high() {
    compile_ok(
        r#"
program test
    integer :: n = 8
    select case (n)
    case (1:5)
        print *, 'low'
    case (6:10)
        print *, 'high'
    end select
end program test
"#,
    );
}

#[test]
fn case_open_upper() {
    compile_ok(
        r#"
program test
    integer :: n = 100
    select case (n)
    case (:0)
        print *, 'non-positive'
    case (1:)
        print *, 'positive'
    end select
end program test
"#,
    );
}

#[test]
fn case_open_lower() {
    compile_ok(
        r#"
program test
    integer :: n = -5
    select case (n)
    case (:0)
        print *, 'non-positive'
    case (1:)
        print *, 'positive'
    end select
end program test
"#,
    );
}

#[test]
fn case_open_both_ends() {
    compile_ok(
        r#"
program test
    integer :: n = 50
    select case (n)
    case (:9)
        print *, 'single digit'
    case (10:99)
        print *, 'double digit'
    case (100:)
        print *, 'triple digit or more'
    end select
end program test
"#,
    );
}

#[test]
fn case_range_boundary_exact() {
    compile_ok(
        r#"
program test
    integer :: n
    do n = 4, 6
        select case (n)
        case (:4)
            print *, 'le 4'
        case (5)
            print *, 'eq 5'
        case (6:)
            print *, 'ge 6'
        end select
    end do
end program test
"#,
    );
}

// ── Multiple values in one case ───────────────────────────────

#[test]
fn case_multiple_values() {
    compile_ok(
        r#"
program test
    integer :: n = 3
    select case (n)
    case (1, 3, 5, 7, 9)
        print *, 'odd'
    case (2, 4, 6, 8, 10)
        print *, 'even'
    end select
end program test
"#,
    );
}

#[test]
fn case_mix_values_and_range() {
    compile_ok(
        r#"
program test
    integer :: n = 0
    select case (n)
    case (0, 1, 2)
        print *, 'small'
    case (3:10)
        print *, 'medium'
    case (11:)
        print *, 'large'
    end select
end program test
"#,
    );
}

#[test]
fn case_multiple_values_some_match() {
    compile_ok(
        r#"
program test
    integer :: i
    do i = 1, 6
        select case (i)
        case (1, 2, 6)
            print *, 'match'
        case default
            print *, 'no'
        end select
    end do
end program test
"#,
    );
}

// ── Character SELECT CASE ─────────────────────────────────────

#[test]
fn case_char_exact() {
    compile_ok(
        r#"
program test
    character :: c = 'b'
    select case (c)
    case ('a')
        print *, 'a'
    case ('b')
        print *, 'b'
    case ('c')
        print *, 'c'
    end select
end program test
"#,
    );
}

#[test]
fn case_char_range() {
    compile_ok(
        r#"
program test
    character :: c = 'm'
    select case (c)
    case ('a':'m')
        print *, 'first half'
    case ('n':'z')
        print *, 'second half'
    end select
end program test
"#,
    );
}

#[test]
fn case_char_open_range() {
    compile_ok(
        r#"
program test
    character :: c = 'Z'
    select case (c)
    case ('A':'Z')
        print *, 'uppercase'
    case ('a':'z')
        print *, 'lowercase'
    case default
        print *, 'other'
    end select
end program test
"#,
    );
}

#[test]
fn case_char_multiple_values() {
    compile_ok(
        r#"
program test
    character :: c = 'e'
    select case (c)
    case ('a', 'e', 'i', 'o', 'u')
        print *, 'vowel'
    case default
        print *, 'consonant'
    end select
end program test
"#,
    );
}

#[test]
fn case_char_string() {
    compile_ok(
        r#"
program test
    character(len=3) :: s = 'foo'
    select case (s)
    case ('bar')
        print *, 'bar'
    case ('baz')
        print *, 'baz'
    case ('foo')
        print *, 'foo'
    case default
        print *, 'other'
    end select
end program test
"#,
    );
}

// ── Nested SELECT CASE ────────────────────────────────────────

#[test]
fn nested_select_case() {
    compile_ok(
        r#"
program test
    integer :: i = 2, j = 3
    select case (i)
    case (1)
        print *, 'i=1'
    case (2)
        select case (j)
        case (1:2)
            print *, 'i=2, j small'
        case (3:)
            print *, 'i=2, j large'
        end select
    case default
        print *, 'other'
    end select
end program test
"#,
    );
}

#[test]
fn nested_select_in_loop() {
    compile_ok(
        r#"
program test
    integer :: i, j
    do i = 1, 3
        select case (i)
        case (1)
            do j = 1, 2
                select case (j)
                case (1)
                    print *, '1,1'
                case (2)
                    print *, '1,2'
                end select
            end do
        case (2:3)
            print *, 'i=', i
        end select
    end do
end program test
"#,
    );
}

// ── SELECT CASE on expression ─────────────────────────────────

#[test]
fn case_on_expression() {
    compile_ok(
        r#"
program test
    integer :: x = 5, y = 3
    select case (x + y)
    case (:7)
        print *, 'small sum'
    case (8:10)
        print *, 'medium sum'
    case (11:)
        print *, 'large sum'
    end select
end program test
"#,
    );
}

#[test]
fn case_on_function_result() {
    compile_ok(
        r#"
program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    select case (sum(a))
    case (:10)
        print *, 'small'
    case (11:20)
        print *, 'medium'
    case (21:)
        print *, 'large'
    end select
end program test
"#,
    );
}

#[test]
fn case_on_mod_result() {
    compile_ok(
        r#"
program test
    integer :: i
    do i = 1, 6
        select case (mod(i, 3))
        case (0)
            print *, i, 'div by 3'
        case (1)
            print *, i, 'rem 1'
        case (2)
            print *, i, 'rem 2'
        end select
    end do
end program test
"#,
    );
}

// ── Default only / fallthrough patterns ───────────────────────

#[test]
fn case_default_only() {
    compile_ok(
        r#"
program test
    integer :: n = 42
    select case (n)
    case default
        print *, 'default'
    end select
end program test
"#,
    );
}

#[test]
fn case_no_match_no_default() {
    compile_ok(
        r#"
program test
    integer :: n = 99
    select case (n)
    case (1)
        print *, 'one'
    case (2)
        print *, 'two'
    end select
    print *, 'after select'
end program test
"#,
    );
}

#[test]
fn case_large_range_integers() {
    compile_ok(
        r#"
program test
    integer :: n = 5000
    select case (n)
    case (1:999)
        print *, 'hundreds'
    case (1000:9999)
        print *, 'thousands'
    case (10000:)
        print *, 'ten-thousands+'
    end select
end program test
"#,
    );
}

// ── Logical SELECT CASE via integer conversion ────────────────

#[test]
fn case_from_logical_merge() {
    compile_ok(
        r#"
program test
    logical :: flag = .true.
    select case (merge(1, 0, flag))
    case (0)
        print *, 'false'
    case (1)
        print *, 'true'
    end select
end program test
"#,
    );
}
