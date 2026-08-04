! vybe-test: fortran/associate_construct_extended/associate_expr_char_concat
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
character(len=8) :: a = 'foo', b = 'bar'
associate (ab => trim(a) // trim(b))
if (trim(ab) /= "foobar") then
    print *, "FAIL: want [foobar] got [", ab, "]"
    stop 1
end if
end associate
end program t
