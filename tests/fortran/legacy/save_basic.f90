! vybe-test: fortran/legacy/save_basic
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    call inc()
    call inc()
    call inc()
contains
    subroutine inc()
        integer, save :: count = 0
        count = count + 1
        print *, count
    end subroutine inc
end program test
