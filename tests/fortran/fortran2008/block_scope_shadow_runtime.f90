! vybe-test: fortran/fortran2008/block_scope_shadow_runtime
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program t
    integer :: i = 10
    block
        integer :: i
        i = 99
        if ((i) /= 99) then
    print *, "FAIL: want [99] got [", i, "]"
    stop 1
end if
    end block
    if ((i) /= 10) then
    print *, "FAIL: want [10] got [", i, "]"
    stop 1
end if
end program t
