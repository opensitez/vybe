! vybe-test: fortran/arrays/array_subroutine_arg
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    call print_first(a)
contains
    subroutine print_first(v)
        integer, intent(in) :: v(:)
        print *, v(1)
    end subroutine
end program test
