! vybe-test: fortran/static_initialization_order/static_initialization_order_character_default_len
! origin: languages/fortran/tests/fortran/test_static_initialization_order.rs

program static_initialization_order_character_default_len
    character(len=4) :: token = 'seed'
    character(len=len_trim(token)) :: short
    short = token
    if (trim(trim(short)) /= "seed") then
    print *, "FAIL: want [seed] got [", trim(short), "]"
    stop 1
end if
    if ((len_trim(short)) /= 4) then
    print *, "FAIL: want [4] got [", len_trim(short), "]"
    stop 1
end if
end program static_initialization_order_character_default_len
