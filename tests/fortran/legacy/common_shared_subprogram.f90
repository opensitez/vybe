! vybe-test: fortran/legacy/common_shared_subprogram
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    integer :: total
    common /result/ total
    total = 0
    call accumulate(5)
    print *, total
contains
    subroutine accumulate(n)
        integer, intent(in) :: n
        integer :: total
        common /result/ total
        total = total + n
    end subroutine accumulate
end program test
