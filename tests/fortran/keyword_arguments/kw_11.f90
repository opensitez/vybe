! vybe-test: fortran/keyword_arguments/kw_11
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
subroutine s(x,y,z,w)
integer::x,y,z,w
end
program p
call s(w=4,z=3,y=2,x=1)
end program p
