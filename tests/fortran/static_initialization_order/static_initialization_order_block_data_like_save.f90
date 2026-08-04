! vybe-test: fortran/static_initialization_order/static_initialization_order_block_data_like_save
! origin: languages/fortran/tests/fortran/test_static_initialization_order.rs

program static_initialization_order_block_data_like_save
    integer, save :: counter
    counter = 5
    if ((counter) /= 5) then
    print *, "FAIL: want [5] got [", counter, "]"
    stop 1
end if
    counter = counter + 1
    if ((counter) /= 6) then
    print *, "FAIL: want [6] got [", counter, "]"
    stop 1
end if
end program static_initialization_order_block_data_like_save
