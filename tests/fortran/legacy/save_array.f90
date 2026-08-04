! vybe-test: fortran/legacy/save_array
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    call store(42)
    call retrieve()
contains
    subroutine store(val)
        integer, intent(in) :: val
        integer, save :: stored
        stored = val
    end subroutine store
    subroutine retrieve()
        integer, save :: stored
        print *, stored
    end subroutine retrieve
end program test
