! vybe-test: fortran/static_initialization_order/static_initialization_order_character_len_from_constant
! origin: languages/fortran/tests/fortran/test_static_initialization_order.rs

program static_initialization_order_character_len_from_constant
    integer, parameter :: name_len = 5
    character(len=name_len) :: tag = 'abc'
    if ((len(tag)) /= 5) then
    print *, "FAIL: want [5] got [", len(tag), "]"
    stop 1
end if
    if (trim(trim(tag)) /= "abc") then
    print *, "FAIL: want [abc] got [", trim(tag), "]"
    stop 1
end if
end program static_initialization_order_character_len_from_constant
