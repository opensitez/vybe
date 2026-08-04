! vybe-test: fortran/logical_eqv_neqv/assign_eqv_result_to_variable_and_print
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
logical :: r
r = .true. .eqv. .false.
if ((r) .neqv. .false.) then
    print *, "FAIL: want [false] got [", r, "]"
    stop 1
end if
end program t
