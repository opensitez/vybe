! vybe-test: fortran/fortran2018_extended/sort_integer_vector_descending
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

program t
    integer :: a(4) = [3, 1, 4, 2]
    call sort(a, reverse=.true.)
    print *, a(1)
end program t
