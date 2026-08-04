! vybe-test: fortran/if_blocks/if_logical_or
! origin: languages/fortran/tests/fortran/test_if_blocks.rs
program t
if (1 > 5 .or. 2 > 1) then
if (trim("either") /= "either") then
    print *, "FAIL: want [either] got [", "either", "]"
    stop 1
end if
end if
end program t
