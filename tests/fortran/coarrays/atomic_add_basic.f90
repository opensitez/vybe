! vybe-test: fortran/coarrays/atomic_add_basic
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    use iso_fortran_env
    integer(atomic_int_kind) :: n[*]
    call atomic_define(n, 0)
    call atomic_add(n, 1)
    print *, n
end program test
