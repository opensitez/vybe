! vybe-test: fortran/if_blocks/if_true
! origin: languages/fortran/tests/fortran/test_if_blocks.rs
program t
if (1 > 0) then
if (trim("yes") /= "yes") then
    print *, "FAIL: want [yes] got [", "yes", "]"
    stop 1
end if
end if
end program t
