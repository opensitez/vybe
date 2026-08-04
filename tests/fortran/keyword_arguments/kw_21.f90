! vybe-test: fortran/keyword_arguments/kw_21
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
subroutine s(x,y,z,w)
integer::x, y, z, w
end
program p
call s(1, 2, w=4, z=3)
end program p
