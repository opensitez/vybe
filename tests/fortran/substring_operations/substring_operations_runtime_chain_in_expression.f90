! vybe-test: fortran/substring_operations/substring_operations_runtime_chain_in_expression
! origin: languages/fortran/tests/fortran/test_substring_operations.rs

program p
character(len=20) :: msg
msg = trim('pre:' // 'fortran' // '-' // 'lang')
if (trim(trim(msg(1:4))) /= "pre:") then
    print *, "FAIL: want [pre:] got [", trim(msg(1:4)), "]"
    stop 1
end if
if (trim(trim(msg(6:10))) /= "ortra") then
    print *, "FAIL: want [ortra] got [", trim(msg(6:10)), "]"
    stop 1
end if
if (trim(trim(msg(12:15))) /= "lang") then
    print *, "FAIL: want [lang] got [", trim(msg(12:15)), "]"
    stop 1
end if
end program p
