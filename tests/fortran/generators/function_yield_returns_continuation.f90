! vybe-test: fortran/generators/function_yield_returns_continuation
! origin: languages/fortran/tests/fortran/test_generators.rs

program test
    if (trim(count()) /= "[continuation]") then
    print *, "FAIL: want [[continuation]] got [", count(), "]"
    stop 1
end if
contains
    function count() result(res)
        yield 1
    end function count
end program test
