! vybe-test: fortran/if_blocks/single_line_if_false_path
! origin: languages/fortran/tests/fortran/test_if_blocks.rs
program t
if (1 == 0) print *, "no"
if (trim("ok") /= "ok") then
    print *, "FAIL: want [ok] got [", "ok", "]"
    stop 1
end if
end program t
