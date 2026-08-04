! vybe-test: fortran/legacy_data_extended/common_shared_accumulator_subprogram
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    integer :: total
    common /acc/ total
    total = 0
    call bump(4)
    call bump(6)
    print *, total
contains
    subroutine bump(n)
        integer, intent(in) :: n
        integer :: total
        common /acc/ total
        total = total + n
    end subroutine bump
end program t
