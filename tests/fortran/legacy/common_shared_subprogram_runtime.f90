! vybe-test: fortran/legacy/common_shared_subprogram_runtime
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    integer :: total
    common /result/ total
    total = 0
    call accumulate(5)
    if ((total) /= 5) then
    print *, "FAIL: want [5] got [", total, "]"
    stop 1
end if
contains
    subroutine accumulate(n)
        integer, intent(in) :: n
        integer :: total
        common /result/ total
        total = total + n
    end subroutine accumulate
end program test
