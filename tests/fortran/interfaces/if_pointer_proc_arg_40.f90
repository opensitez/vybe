! vybe-test: fortran/interfaces/if_pointer_proc_arg_40
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
integer :: hits = 0
contains
subroutine target_sub()
hits = hits + 1
end subroutine target_sub
subroutine apply(f)
procedure() :: f
call f()
end subroutine apply
end module m
program t
use m
call apply(target_sub)
if (hits /= 1) then
    print *, "FAIL: want [1] got [", hits, "]"
    stop 1
end if
end program t
