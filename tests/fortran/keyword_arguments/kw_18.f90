! vybe-test: fortran/keyword_arguments/kw_18
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
subroutine s(x,y)
integer, intent(in)::x,y
end
program p
call s(y=2,x=1)
end program p
