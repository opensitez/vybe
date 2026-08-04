! vybe-test: fortran/attributes/attr_save_runtime_default_stable
! origin: languages/fortran/tests/fortran/test_attributes.rs

program attr_save_runtime_default_stable
    integer, save :: counter = 0
    counter = counter + 1
    if ((counter) /= 1) then
    print *, "FAIL: want [1] got [", counter, "]"
    stop 1
end if
    counter = counter + 1
    if ((counter) /= 2) then
    print *, "FAIL: want [2] got [", counter, "]"
    stop 1
end if
end program attr_save_runtime_default_stable
