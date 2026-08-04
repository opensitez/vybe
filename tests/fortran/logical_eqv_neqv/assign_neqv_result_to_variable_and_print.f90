! vybe-test: fortran/logical_eqv_neqv/assign_neqv_result_to_variable_and_print
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
logical :: r
r = .false. .neqv. .true.
if ((r) .neqv. .true.) then
    print *, "FAIL: want [true] got [", r, "]"
    stop 1
end if
end program t
