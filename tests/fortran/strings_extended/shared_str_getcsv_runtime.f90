! vybe-test: fortran/strings_extended/shared_str_getcsv_runtime
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=256), allocatable :: fields(:)
fields = str_getcsv('"Smith, John",42,"New York","Engineer, Senior",95000.50')
if ((size(fields)) /= 5) then
    print *, "FAIL: want [5] got [", size(fields), "]"
    stop 1
end if
if (trim(trim(fields(1))) /= "Smith, John") then
    print *, "FAIL: want [Smith, John] got [", trim(fields(1)), "]"
    stop 1
end if
if (trim(trim(fields(4))) /= "Engineer, Senior") then
    print *, "FAIL: want [Engineer, Senior] got [", trim(fields(4)), "]"
    stop 1
end if
if (abs((trim(fields(5))) - 95000.50) > 1.0e-6) then
    print *, "FAIL: want [95000.50] got [", trim(fields(5)), "]"
    stop 1
end if
end program t
