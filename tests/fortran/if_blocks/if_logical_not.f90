! vybe-test: fortran/if_blocks/if_logical_not
! origin: languages/fortran/tests/fortran/test_if_blocks.rs
program t
if (.not. (1 > 5)) then
if (trim("negated") /= "negated") then
    print *, "FAIL: want [negated] got [", "negated", "]"
    stop 1
end if
end if
end program t
