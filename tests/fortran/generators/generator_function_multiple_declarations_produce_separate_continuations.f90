! vybe-test: fortran/generators/generator_function_multiple_declarations_produce_separate_continuations
! origin: languages/fortran/tests/fortran/test_generators.rs

program test
    if (trim(make()) /= "[continuation]") then
    print *, "FAIL: want [[continuation]] got [", make(), "]"
    stop 1
end if
    if (trim(make()) /= "[continuation]") then
    print *, "FAIL: want [[continuation]] got [", make(), "]"
    stop 1
end if
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
