! vybe-test: fortran/kinds/kind_array
! origin: languages/fortran/tests/fortran/test_kinds.rs

program test
    integer(kind=8) :: a(3) = [1_8, 2_8, 3_8]
    print *, a(1)
end program test
