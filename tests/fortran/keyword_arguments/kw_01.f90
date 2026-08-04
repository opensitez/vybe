! vybe-test: fortran/keyword_arguments/kw_01
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
subroutine s(x,y)
integer::x,y
end
program p
call s(x=1,y=2)
end program p
