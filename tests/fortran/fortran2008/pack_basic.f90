! vybe-test: fortran/fortran2008/pack_basic
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    logical :: mask(5) = [.true., .false., .true., .false., .true.]
    integer :: b(3)
    b = pack(a, mask)
    print *, b(1)
end program test
