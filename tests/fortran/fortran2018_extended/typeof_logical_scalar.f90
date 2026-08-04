! vybe-test: fortran/fortran2018_extended/typeof_logical_scalar
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

program t
    logical :: flag = .true.
    print *, typeof(flag)
end program t
