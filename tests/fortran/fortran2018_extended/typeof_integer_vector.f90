! vybe-test: fortran/fortran2018_extended/typeof_integer_vector
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

program t
    integer :: v(3) = [1, 2, 3]
    print *, typeof(v)
end program t
