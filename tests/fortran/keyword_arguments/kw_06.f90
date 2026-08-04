! vybe-test: fortran/keyword_arguments/kw_06
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
subroutine s(x,y)
integer, optional::x,y
end
program p
call s(x=1)
end program p
