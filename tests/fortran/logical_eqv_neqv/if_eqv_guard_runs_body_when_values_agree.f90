! vybe-test: fortran/logical_eqv_neqv/if_eqv_guard_runs_body_when_values_agree
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if (.false. .eqv. .false.) then
if (trim("run") /= "run") then
    print *, "FAIL: want [run] got [", "run", "]"
    stop 1
end if
end if
if (trim("done") /= "done") then
    print *, "FAIL: want [done] got [", "done", "]"
    stop 1
end if
end program t
