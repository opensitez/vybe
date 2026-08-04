! vybe-test: fortran/coarrays/atomic_fetch_add
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    use iso_fortran_env
    integer(atomic_int_kind) :: counter[*]
    integer :: prev
    call atomic_define(counter, 5)
    call atomic_fetch_add(counter, 3, prev)
    print *, prev
    print *, counter
end program test
