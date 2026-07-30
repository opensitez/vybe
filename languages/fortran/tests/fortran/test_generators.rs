use super::helpers::run_prints;

#[test]
fn function_yield_returns_continuation() {
    let out = run_prints(
        r#"
program test
    print *, count()
contains
    function count() result(res)
        yield 1
    end function count
end program test
"#,
    );
    assert_eq!(out, vec!["[continuation]"]);
}

#[test]
fn generator_function_is_lazy() {
    let out = run_prints(
        r#"
program test
    print *, 100
    print *, count()
    print *, 200
contains
    function count() result(res)
        integer :: n
        print *, "should-not-run"
        n = 1
        yield n
    end function count
end program test
"#,
    );
    assert_eq!(out, vec!["100", "[continuation]", "200"]);
}

#[test]
fn generator_function_multiple_declarations_produce_separate_continuations() {
    let out = run_prints(
        r#"
program test
    print *, make()
    print *, make()
contains
    function make() result(res)
        integer :: n
        n = 1
        if (n > 0) then
            n = 1
            yield n
        else
            n = 2
            yield n
        end if
        n = 3
        yield n
    end function make
end program test
"#,
    );
    assert_eq!(out, vec!["[continuation]", "[continuation]"]);
}
