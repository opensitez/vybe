! vybe-test: fortran/coarrays/atomic_define_basic
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    use iso_fortran_env
    integer(atomic_int_kind) :: counter[*]
    call atomic_define(counter, 0)
    print *, counter
end program test
