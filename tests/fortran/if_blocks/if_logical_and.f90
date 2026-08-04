! vybe-test: fortran/if_blocks/if_logical_and
! origin: languages/fortran/tests/fortran/test_if_blocks.rs
program t
if (1 > 0 .and. 2 > 1) then
if (trim("both") /= "both") then
    print *, "FAIL: want [both] got [", "both", "]"
    stop 1
end if
end if
end program t
