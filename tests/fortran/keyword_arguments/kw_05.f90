! vybe-test: fortran/keyword_arguments/kw_05
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
subroutine s(x,y)
integer, optional::x,y
end
program p
call s(y=2)
end program p
