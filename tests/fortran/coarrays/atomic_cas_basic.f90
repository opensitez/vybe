! vybe-test: fortran/coarrays/atomic_cas_basic
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    use iso_fortran_env
    integer(atomic_int_kind) :: x[*]
    integer :: old
    call atomic_define(x, 10)
    call atomic_cas(x, old, 10, 20)
    print *, x
end program test
