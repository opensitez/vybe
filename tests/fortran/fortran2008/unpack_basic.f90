! vybe-test: fortran/fortran2008/unpack_basic
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program test
    integer :: a(3) = [10, 20, 30]
    logical :: mask(5) = [.true., .false., .true., .false., .true.]
    integer :: b(5)
    integer :: fill(5) = [0, 0, 0, 0, 0]
    b = unpack(a, mask, fill)
    print *, b(1)
end program test
