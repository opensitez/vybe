! vybe-test: fortran/array_semantics/arr_unpack_19
! origin: languages/fortran/tests/fortran/test_array_semantics.rs
program p
integer::a(2)=[1,2],f(3)=[0,0,0]
print *,unpack(a,[.true.,.false.,.true.],f)
end program p
