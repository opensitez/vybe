! vybe-test: fortran/logical_eqv_neqv/neqv_via_logical_variables_true_false
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
logical :: a = .true., b = .false.
if ((a .neqv. b) .neqv. .true.) then
    print *, "FAIL: want [true] got [", a .neqv. b, "]"
    stop 1
end if
end program t
