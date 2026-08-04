! vybe-test: fortran/fortran2008/contiguous_dummy
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    call process(a)
contains
    subroutine process(v)
        integer, intent(in), contiguous :: v(:)
        print *, v(1)
    end subroutine process
end program test
