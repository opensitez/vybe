! vybe-test: fortran/kinds/selected_int_kind_9
! origin: languages/fortran/tests/fortran/test_kinds.rs

program test
    integer, parameter :: k = selected_int_kind(9)
    integer(kind=k) :: n = 999999999
    print *, n
end program test
