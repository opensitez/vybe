! vybe-test: fortran/contiguous_attributes_and_checks/test_contiguous_attributes_and_checks_valid_contiguous_section
! origin: languages/fortran/tests/fortran/test_contiguous_attributes_and_checks.rs

program test_contiguous_attributes_and_checks
    implicit none
    real :: values(5)
    values = (/1.0, 2.0, 3.0, 4.0, 5.0/)
    call inspect(values)

contains
    subroutine inspect(a)
        real, contiguous, intent(in) :: a(:)
        if ((size(a)) /= 5) then
    print *, "FAIL: want [5] got [", size(a), "]"
    stop 1
end if
    end subroutine inspect
end program test_contiguous_attributes_and_checks
