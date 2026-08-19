! vybe-test: fortran/interface_blocks/interface_blocks_05
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
program p
interface
subroutine s(x)
integer::x
end subroutine s
end interface
call s(1)
end program p

subroutine s(x)
integer::x
end subroutine s
