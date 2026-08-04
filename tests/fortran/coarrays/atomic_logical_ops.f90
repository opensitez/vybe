! vybe-test: fortran/coarrays/atomic_logical_ops
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    use iso_fortran_env
    logical(atomic_logical_kind) :: flag[*]
    call atomic_define(flag, .false.)
    call atomic_or(flag, .true.)
    print *, flag
end program test
