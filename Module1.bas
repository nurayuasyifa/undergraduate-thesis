Attribute VB_Name = "Module1"
Global con As New ADODB.Connection
Global rs As ADODB.Recordset
Global dataasli(1000, 10) As Double
Global datanorm(1000, 10) As Double
Global dataasliprediksi(1000, 10) As Double
Global solusirk(1000, 100, 3) As Double
Global eror(1000, 100, 4) As Double
Global erorprediksi(100, 4) As Double
Global normalisasi(100, 4) As Double
Global deltafix As Double, omegafix As Double, miufix As Double, epsilonfix As Double, tetafix As Double, gurufix As Double
Global nofix As Integer, maxitfix As Integer, nfix As Integer, mmrefix As Double
Global mmreakhir As Double, mmres As Double, mmrei As Double, mmrer As Double, mmreakhirp As Double
Global maks(4) As Double, mini(4) As Double
Global poladata(3, 100, 2) As Double
Global poladataprediksi(3, 100, 2) As Double, poladatauji(3, 36, 2) As Double
Global target(3, 100) As Double
Global targetprediksi(3, 100) As Double, targetuji(3, 36) As Double
Global mu As Double, beta As Byte, epoch As Integer, err As Double, m As Double, epochakhir As Integer, mseakhir As Double
Global bbv(3, 2, 2) As Double, bbw(3, 2, 2) As Double, bbvn(3, 2, 2) As Double, bbwn(3, 2, 1) As Double
Global bbvp(3, 2, 2) As Double, bbwp(3, 2, 2) As Double, bbvnp(3, 2, 2) As Double, bbwnp(3, 2, 1) As Double
Global zin(3, 100, 3) As Double, z(3, 100, 3) As Double, yin(3, 100, 3) As Double, Y(3, 100, 3) As Double
Global zinp(3, 100, 3) As Double, zp(3, 100, 3) As Double, yinp(3, 100, 3) As Double, yp(3, 100, 3) As Double
Global zinpu(3, 100, 3) As Double, zpu(3, 100, 3) As Double, yinpu(3, 100, 3) As Double, ypu(3, 100, 3) As Double
Global outputs(3, 100) As Double, outputfix(3, 100) As Double, outputsp(3, 100) As Double, outputfixp(3, 100) As Double
Global outputspu(3, 100) As Double, outputfixpu(3, 100) As Double
Global denormalisasi(100, 3) As Double, denormalisasiprediksi(100, 3) As Double
Global mmreprediksi As Double, mmresprediksi As Double, mmreiprediksi As Double, mmrerprediksi As Double
'menghitung mmre
Function mre(asli As Double, estimasi As Double) As Double
mre = Abs(asli - estimasi) / asli
End Function
'menghitung mae
Function mae(asli As Double, estimasi As Double) As Double
mae = Abs(asli - estimasi)
End Function
'menghitung mse
Function msse(asli As Double, estimasi As Double) As Double
msse = (asli - estimasi) ^ 2
End Function
'random angka -1 hingga 1
Function randomantara(lowerbound As Double, upperbound As Double) As Double
    Randomize
    randomantara = ((upperbound - lowerbound) * Rnd + lowerbound)
End Function
'fungsi sigmoid bipolar
Function sigmoidbipolar(X As Double) As Double
sigmoidbipolar = (1 - Exp(-2 * X)) / (1 + Exp(-2 * X))
End Function
Function sigmoidbiner(X As Double) As Double
sigmoidbiner = 1 / (1 + Exp(-1 * X))
End Function
'turunan fungsi sigmoid bipolar
Function slope(X As Double) As Double
slope = (1 + X) * (1 - X)
End Function
