! vybe-test: fortran/keyword_arguments/kw_02
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
subroutine s(x,y)
integer::x,y
end
program p
call s(y=2,x=1)
end program p
