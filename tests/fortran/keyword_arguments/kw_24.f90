! vybe-test: fortran/keyword_arguments/kw_24
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
subroutine s(x,y,z)
integer, optional :: x, y, z
end
program p
call s(x=1, z=3)
end program p
