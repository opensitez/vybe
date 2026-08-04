! vybe-test: fortran/strings_extended/shared_str_getcsv_direct_runtime
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
if ((size(str_getcsv('"Smith, John",42,"New York","Engineer, Senior",95000.50'))) /= 5) then
    print *, "FAIL: want [5] got [", size(str_getcsv('"Smith, John",42,"New York","Engineer, Senior",95000.50')), "]"
    stop 1
end if
end program t
