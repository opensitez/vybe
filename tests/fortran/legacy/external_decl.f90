! vybe-test: fortran/legacy/external_decl
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    external :: my_func
    print *, "ok"
contains
    function my_func(x)
        real :: my_func, x
        my_func = x * 2.0
    end function my_func
end program test
