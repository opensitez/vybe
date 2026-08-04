! vybe-test: fortran/allocate_statement/allocate_statement_runtime_character_alloc_and_length
! origin: languages/fortran/tests/fortran/test_allocate_statement.rs
program t
character(len=:), allocatable :: s
allocate(character(len=4) :: s)
s = 'fort'
if ((len(s)) /= 4) then
    print *, "FAIL: want [4] got [", len(s), "]"
    stop 1
end if
if (trim(s) /= "fort") then
    print *, "FAIL: want [fort] got [", s, "]"
    stop 1
end if
end program t
