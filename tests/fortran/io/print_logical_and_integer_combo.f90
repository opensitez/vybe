! vybe-test: fortran/io/print_logical_and_integer_combo
! origin: languages/fortran/tests/fortran/test_io.rs

program test
    if (trim("alive=") /= "alive=") then
    print *, "FAIL: want [alive=] got [", "alive=", "]"
    stop 1
end if
    if ((.true.) .neqv. .true.) then
    print *, "FAIL: want [true] got [", .true., "]"
    stop 1
end if
    if (trim("count=") /= "count=") then
    print *, "FAIL: want [count=] got [", "count=", "]"
    stop 1
end if
    if ((3) /= 3) then
    print *, "FAIL: want [3] got [", 3, "]"
    stop 1
end if
end program test
