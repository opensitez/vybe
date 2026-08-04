! vybe-test: fortran/generators/generator_function_is_lazy
! origin: languages/fortran/tests/fortran/test_generators.rs

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
