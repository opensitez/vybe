! vybe-test: fortran/alternate_returns/alternate_returns_runtime_plain_return_defaults_to_call_fallthrough
! origin: languages/fortran/tests/fortran/test_alternate_returns.rs
program p
integer :: x
x = 0
call s(*10,*20)
x = 1
if ((x) /= 1) then
    print *, "FAIL: want [1] got [", x, "]"
    stop 1
end if
10 print *, x
20 continue
end program p
subroutine s(*,*)
return
end
