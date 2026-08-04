! vybe-test: fortran/coarrays/coarray_logical_decl
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    logical :: flag[*]
    flag = .true.
    print *, flag
end program test
