! vybe-test: fortran/if_construct_extended/if_no_else_true_branch_runs
! origin: languages/fortran/tests/fortran/test_if_construct_extended.rs
program t
if (7 > 3) then
if (trim("ran") /= "ran") then
    print *, "FAIL: want [ran] got [", "ran", "]"
    stop 1
end if
end if
end program t
