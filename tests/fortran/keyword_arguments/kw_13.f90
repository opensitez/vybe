! vybe-test: fortran/keyword_arguments/kw_13
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
subroutine s(flag,val)
logical::flag
integer::val
end
program p
call s(val=2, flag=.true.)
end program p
