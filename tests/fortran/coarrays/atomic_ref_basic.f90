! vybe-test: fortran/coarrays/atomic_ref_basic
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    use iso_fortran_env
    integer(atomic_int_kind) :: counter[*]
    integer :: val
    call atomic_define(counter, 7)
    call atomic_ref(val, counter)
    print *, val
end program test
