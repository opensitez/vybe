! vybe-test: fortran/fortran2018_extended/sort_integer_vector_ascending
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

program t
    integer :: a(5) = [3, 1, 4, 1, 5]
    call sort(a)
    print *, a(1), a(5)
end program t
