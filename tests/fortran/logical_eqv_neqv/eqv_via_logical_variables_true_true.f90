! vybe-test: fortran/logical_eqv_neqv/eqv_via_logical_variables_true_true
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
logical :: a = .true., b = .true.
if ((a .eqv. b) .neqv. .true.) then
    print *, "FAIL: want [true] got [", a .eqv. b, "]"
    stop 1
end if
end program t
