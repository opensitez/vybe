! vybe-test: fortran/keyword_arguments/kw_20
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
subroutine s(x,y,z)
integer::x,y,z
end
program p
call s(x=1,z=3,y=2)
end program p
