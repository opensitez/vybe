! vybe-test: fortran/if_construct_extended/if_single_line_if_false_path
! origin: languages/fortran/tests/fortran/test_if_construct_extended.rs
program t
if (1 == 2) print *, 'single-false'
if (trim('after') /= "after") then
    print *, "FAIL: want [after] got [", 'after', "]"
    stop 1
end if
end program t
