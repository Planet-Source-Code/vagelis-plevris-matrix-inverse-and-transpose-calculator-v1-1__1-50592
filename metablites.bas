Attribute VB_Name = "Variables"
'Define Global variables

'Matrix dimensions are set to Max 10x10 for the interface needs, but can be increased here to whatever
Global Const MAX_DIM = 10

Global System_DIM As Integer 'Current Matrix [A] dimensions
Global Matrix_A(1 To MAX_DIM, 1 To MAX_DIM)
Global Operations_Matrix(1 To MAX_DIM, 1 To 2 * MAX_DIM) 'Matrix where the calculations are done
Global Inverse_Matrix(1 To MAX_DIM, 1 To MAX_DIM) 'Matrix with the Inverse of [A]
Global Transpose_Matrix(1 To MAX_DIM, 1 To MAX_DIM) 'Matrix with the Transpose of [A]
Global Solution_Problem As Boolean 'Determines whether the inverse was found or not
Global Matrix_Mult(1 To MAX_DIM, 1 To MAX_DIM) 'Matrix with the product [A]*[A-1]=[I] (must be always equal to Singular matrix [I])

