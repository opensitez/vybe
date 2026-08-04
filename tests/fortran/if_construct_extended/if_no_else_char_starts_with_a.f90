! vybe-test: fortran/if_construct_extended/if_no_else_char_starts_with_a
! origin: languages/fortran/tests/fortran/test_if_construct_extended.rs
program t
character(len=5) :: word = "alpha"
if (word(1:1) == "a") then
if (trim("a-word") /= "a-word") then
    print *, "FAIL: want [a-word] got [", "a-word", "]"
    stop 1
end if
end if
end program t
