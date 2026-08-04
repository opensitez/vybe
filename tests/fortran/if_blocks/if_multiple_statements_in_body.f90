! vybe-test: fortran/if_blocks/if_multiple_statements_in_body
! origin: languages/fortran/tests/fortran/test_if_blocks.rs
program t
if (1 > 0) then
if (trim("a") /= "a") then
    print *, "FAIL: want [a] got [", "a", "]"
    stop 1
end if
if (trim("b") /= "b") then
    print *, "FAIL: want [b] got [", "b", "]"
    stop 1
end if
if (trim("c") /= "c") then
    print *, "FAIL: want [c] got [", "c", "]"
    stop 1
end if
end if
end program t
