! vybe-test: fortran/logical_eqv_neqv/neqv_with_not_keyword_case_mix
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if ((.FALSE. .neqv. .NOT. .TRUE.) .neqv. .true.) then
    print *, "FAIL: want [true] got [", .FALSE. .neqv. .NOT. .TRUE., "]"
    stop 1
end if
end program t
