! vybe-test: fortran/legacy/save_array_runtime
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
        if ((stored) /= 0) then
    print *, "FAIL: want [0] got [", stored, "]"
    stop 1
end if
    end subroutine retrieve
end program test
