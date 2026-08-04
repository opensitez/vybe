! vybe-test: fortran/move_alloc/move_alloc_04
! origin: languages/fortran/tests/fortran/test_move_alloc.rs
program p
character(len=:), allocatable :: a,b
allocate(character(len=3) :: a)
call move_alloc(a,b)
end program p
