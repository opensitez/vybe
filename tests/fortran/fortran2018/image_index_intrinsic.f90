! vybe-test: fortran/fortran2018/image_index_intrinsic
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    integer :: idx
    integer :: sub(2) = [1, 1]
    idx = image_index([2,2], sub)
    print *, idx
end program test
