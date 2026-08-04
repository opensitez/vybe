! vybe-test: fortran/keyword_arguments/kw_14
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
subroutine s(x,y)
character(len=*)::x
automatic integer::y
end
program p
call s(y=1, x='a')
end program p
