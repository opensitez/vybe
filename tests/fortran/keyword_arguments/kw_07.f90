! vybe-test: fortran/keyword_arguments/kw_07
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
subroutine s(a,b,c)
real::a,b,c
end
program p
call s(c=3.0,b=2.0,a=1.0)
end program p
