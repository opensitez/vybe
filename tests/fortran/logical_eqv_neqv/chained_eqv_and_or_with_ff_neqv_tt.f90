! vybe-test: fortran/logical_eqv_neqv/chained_eqv_and_or_with_ff_neqv_tt
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if (((.false. .eqv. .false.) .and. (.true. .neqv. .true.)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", (.false. .eqv. .false.) .and. (.true. .neqv. .true.), "]"
    stop 1
end if
end program t
