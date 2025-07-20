' 曲线工具集
'                       -- 20231206 小豆

' ==================================================
' 使用简述:
' 计算前先 Call CUSetP0(...) ~ CUSetP3(...) 设置好各个控制点的位置
' 而后运行 CUCalcBezier__(t) 函数获得计算结果, 其中 t 取 0~1

' ==================================================
' 内建变量

' 顶点数据
Dim libCurver_p0_x As Double = 0
Dim libCurver_p0_y As Double = 0
Dim libCurver_p1_x As Double = 0
Dim libCurver_p1_y As Double = 0
Dim libCurver_p2_x As Double = 0
Dim libCurver_p2_y As Double = 0
Dim libCurver_p3_x As Double = 0
Dim libCurver_p3_y As Double = 0

' 长度数据
Dim libCurver_bezier_lenght3 As Double = -1
Dim libCurver_bezier_lenght2 As Double = -1
Dim libCurver_bezier_lenght1 As Double = -1

' 临时变量
Dim libCurver_tempA As Double = 0           ' 临时变量 A
Dim libCurver_tempB As Double = 0           ' 临时变量 B
Dim libCurver_tempC As Byte = 0             ' 临时变量 C
Dim libCurver_tempD As Double = 0           ' 临时变量 D
Dim libCurver_tempE As Double = 0           ' 临时变量 E
Dim libCurver_tempF As Double = 0           ' 临时变量 F
Dim libCurver_tempG As Double = 0           ' 临时变量 G
Dim libCurver_t As Double = 0               ' t
Dim libCurver_t2 As Double = 0              ' t^2
Dim libCurver_t3 As Double = 0              ' t^3
Dim libCurver_n_t As Double = 1             ' 1 - t
Dim libCurver_n_t2 As Double = 1            ' (1 - t)^2
Dim libCurver_n_t3 As Double = 1            ' (1 - t)^3
Dim libCurver_timeStampSplitCnt As Long = 1 ' 时间戳分割
Dim libCurver_timeStampStart As Double = 0  ' 时间戳
Dim libCurver_timeStampEnd As Double = 0    ' 时间戳

' ==================================================
' 设置各个控制点的参数

' 设置 0 点
Export Script CUSetP0(x As Double, y As Double)
    If Abs(x - libCurver_p0_x) >= 0.01 Or Abs(y - libCurver_p0_y) >= 0.01 Then
        libCurver_bezier_lenght1 = -1
        libCurver_bezier_lenght2 = -1
        libCurver_bezier_lenght3 = -1
        libCurver_p0_x = x
        libCurver_p0_y = y
    End If
End Script

' 设置 1 点
Export Script CUSetP1(x As Double, y As Double)
    If Abs(x - libCurver_p1_x) >= 0.01 Or Abs(y - libCurver_p1_y) >= 0.01 Then
        libCurver_bezier_lenght1 = -1
        libCurver_bezier_lenght2 = -1
        libCurver_bezier_lenght3 = -1
        libCurver_p1_x = x
        libCurver_p1_y = y
    End If
End Script

' 设置 2 点
Export Script CUSetP2(x As Double, y As Double)
    If Abs(x - libCurver_p2_x) >= 0.01 Or Abs(y - libCurver_p2_y) >= 0.01 Then
        libCurver_bezier_lenght2 = -1
        libCurver_bezier_lenght3 = -1
        libCurver_p2_x = x
        libCurver_p2_y = y
    End If
End Script

' 设置 3 点
Export Script CUSetP3(x As Double, y As Double)
    If Abs(x - libCurver_p3_x) >= 0.01 Or Abs(y - libCurver_p3_y) >= 0.01 Then
        libCurver_bezier_lenght3 = -1
        libCurver_p3_x = x
        libCurver_p3_y = y
    End If
End Script

' ==================================================
' 计算方法

' ------------------------------------
' ----------------------- 内建方法 BEGIN

Script CUInner_SetT(t As Double)
    If Abs(t - libCurver_t) >= 0.000000001 Then
        libCurver_t = t
        libCurver_t2 = t * t
        libCurver_t3 = libCurver_t2 * t

        libCurver_n_t = 1 - t
        libCurver_n_t2 = libCurver_n_t * libCurver_n_t
        libCurver_n_t3 = libCurver_n_t2 * libCurver_n_t
    End If
End Script

Script CUInner_CalcBezier3(p0 As Double, p1 As Double, p2 As Double, p3 As Double, Return Double)
    Return libCurver_n_t3 * p0 + 3 * libCurver_n_t2 * libCurver_t * p1 + 3 * libCurver_n_t * libCurver_t2 * p2 + libCurver_t3 * p3
End Script

Script CUInner_CalcBezier2(p0 As Double, p1 As Double, p2 As Double, Return Double)
    Return libCurver_n_t2 * p0 + 2 * libCurver_n_t * libCurver_t * p1 + libCurver_t2 * p2
End Script

Script CUInner_CalcBezier1(p0 As Double, p1 As Double, Return Double)
    Return libCurver_n_t2 * p0 + libCurver_t2 * p1
End Script

Script CUInner_CalcDistancePP(x1 As Double, y1 As Double, x2 As Double, y2 As Double, Return Double)
    x1 = x1 - x2
    y1 = y1 - y2
    Return Sqr(x1 * x1 + y1 * y1)
End Script

' ----------------------- 内建方法 END
' ------------------------------------

' ------------------------------------
' ----------------------- 计算贝塞尔曲线结果坐标 BEGIN

' 计算三次曲线 x 坐标
Export Script CUCalcBezier3X(t As Double, Return Double)
    Call CUInner_SetT(t)
    Return CUInner_CalcBezier3(libCurver_p0_x, libCurver_p1_x, libCurver_p2_x, libCurver_p3_x)
End Script

' 计算三次曲线 y 坐标
Export Script CUCalcBezier3Y(t As Double, Return Double)
    Call CUInner_SetT(t)
    Return CUInner_CalcBezier3(libCurver_p0_y, libCurver_p1_y, libCurver_p2_y, libCurver_p3_y)
End Script

' 计算二次曲线 x 坐标
Export Script CUCalcBezier2X(t As Double, Return Double)
    Call CUInner_SetT(t)
    Return CUInner_CalcBezier2(libCurver_p0_x, libCurver_p1_x, libCurver_p2_x)
End Script

' 计算二次曲线 y 坐标
Export Script CUCalcBezier2Y(t As Double, Return Double)
    Call CUInner_SetT(t)
    Return CUInner_CalcBezier2(libCurver_p0_y, libCurver_p1_y, libCurver_p2_y)
End Script

' 计算一次曲线 x 坐标
Export Script CUCalcBezier1X(t As Double, Return Double)
    Call CUInner_SetT(t)
    Return CUInner_CalcBezier1(libCurver_p0_x, libCurver_p1_x)
End Script

' 计算一次曲线 y 坐标
Export Script CUCalcBezier1Y(t As Double, Return Double)
    Call CUInner_SetT(t)
    Return CUInner_CalcBezier1(libCurver_p0_y, libCurver_p1_y)
End Script

' ----------------------- 计算贝塞尔曲线结果坐标 END
' ------------------------------------

' ------------------------------------
' ----------------------- 计算贝塞尔曲线长度 BEGIN

' 近似计算三阶贝塞尔曲线长度
Export Script CUCalcBezier3Len(Return Double)
    If libCurver_bezier_lenght3 > 0 Then
        Return libCurver_bezier_lenght3
    End If

    libCurver_tempB = 0
    libCurver_tempD = libCurver_p0_x
    libCurver_tempE = libCurver_p0_y
    For libCurver_tempC = 1 To 25
        libCurver_tempA = libCurver_tempC / 25

        libCurver_tempF = CUCalcBezier3X(libCurver_tempA)
        libCurver_tempG = CUCalcBezier3Y(libCurver_tempA)
        libCurver_tempB += CUInner_CalcDistancePP(libCurver_tempD, libCurver_tempE, libCurver_tempF, libCurver_tempG)
        libCurver_tempD = libCurver_tempF
        libCurver_tempE = libCurver_tempG
    Next
    libCurver_bezier_lenght3 = libCurver_tempB
    Return libCurver_bezier_lenght3
End Script

' 近似计算二阶贝塞尔曲线长度
Export Script CUCalcBezier2Len(Return Double)
    If libCurver_bezier_lenght2 > 0 Then
        Return libCurver_bezier_lenght2
    End If

    libCurver_tempB = 0
    libCurver_tempD = libCurver_p0_x
    libCurver_tempE = libCurver_p0_y
    For libCurver_tempC = 1 To 25
        libCurver_tempA = libCurver_tempC / 25

        libCurver_tempF = CUCalcBezier2X(libCurver_tempA)
        libCurver_tempG = CUCalcBezier2Y(libCurver_tempA)
        libCurver_tempB += CUInner_CalcDistancePP(libCurver_tempD, libCurver_tempE, libCurver_tempF, libCurver_tempG)
        libCurver_tempD = libCurver_tempF
        libCurver_tempE = libCurver_tempG
    Next
    libCurver_bezier_lenght2 = libCurver_tempB
    Return libCurver_bezier_lenght2
End Script

' 计算一阶贝塞尔曲线长度
Export Script CUCalcBezier1Len(Return Double)
    If libCurver_bezier_lenght1 > 0 Then
        Return libCurver_bezier_lenght1
    End If

    libCurver_bezier_lenght1 = CUInner_CalcDistancePP(libCurver_p0_x, libCurver_p0_y, libCurver_p1_x, libCurver_p1_y)
    Return libCurver_bezier_lenght1
End Script

' ----------------------- 计算贝塞尔曲线长度 END
' ------------------------------------

' ------------------------------------
' ----------------------- 应用曲线 BEGIN

' 应用曲线到 bitmap 位置坐标
' params id bitmap id
' params t 曲线参数
Export Script CUSetBmpPosBezier3(id As Long, t As Double)
    Bitmap(id).destx = CUCalcBezier3X(t)
    Bitmap(id).desty = CUCalcBezier3Y(t)
End Script

' 应用曲线到 bitmap 位置坐标
' params id bitmap id
' params t 曲线参数
Export Script CUSetBmpPosBezier2(id As Long, t As Double)
    Bitmap(id).destx = CUCalcBezier2X(t)
    Bitmap(id).desty = CUCalcBezier2Y(t)
End Script

' 应用曲线到 bitmap 位置坐标
' params id bitmap id
' params t 曲线参数
Export Script CUSetBmpPosBezier1(id As Long, t As Double)
    Bitmap(id).destx = CUCalcBezier1X(t)
    Bitmap(id).desty = CUCalcBezier1Y(t)
End Script

' 应用曲线到 bitmap 缩放
' params id bitmap id
' params t 曲线参数
Export Script CUSetBmpScaleBezier3(id As Long, t As Double)
    Bitmap(id).scalex = CUCalcBezier3X(t)
    Bitmap(id).scaley = CUCalcBezier3Y(t)
End Script

' 应用曲线到 bitmap 缩放
' params id bitmap id
' params t 曲线参数
Export Script CUSetBmpScaleBezier2(id As Long, t As Double)
    Bitmap(id).scalex = CUCalcBezier2X(t)
    Bitmap(id).scaley = CUCalcBezier2Y(t)
End Script

' 应用曲线到 bitmap 缩放
' params id bitmap id
' params t 曲线参数
Export Script CUSetBmpScaleBezier1(id As Long, t As Double)
    Bitmap(id).scalex = CUCalcBezier1X(t)
    Bitmap(id).scaley = CUCalcBezier1Y(t)
End Script

' ----------------------- 应用曲线 END
' ------------------------------------

' ------------------------------------
' ----------------------- 时间工具集 BEGIN

' 设置时间戳端点
' @params startTimeStamp 时间戳起点
' @params endTimeStamp 时间戳终点
Export Script CUTimeSetStamp(startTimeStamp As Double, endTimeStamp As Double)
    libCurver_timeStampEnd = endTimeStamp
    libCurver_timeStampStart = startTimeStamp
End Script

' 设置时间戳分割数
' @params timeStampSplit 时间戳分割数
Export Script CUTimeSetSplit(timeStampSplit As Long)
    If timeStampSplit <= 0 Then
        libCurver_timeStampSplitCnt = 1
    Else
        libCurver_timeStampSplitCnt = timeStampSplit
    End If
End Script

' 计算 t 参数
' @params timeStamp 当前时间戳
' @return t 参数
Export Script CUTimeCalcT(timeStamp As Double, Return Double)
    If Abs(libCurver_timeStampEnd - libCurver_timeStampStart) < 0.000000001 Then
        Return 1
    End If

    timeStamp = (timeStamp - libCurver_timeStampStart) / (libCurver_timeStampEnd - libCurver_timeStampStart)
    If timeStamp < 0 Then
        Return 0
    ElseIf timeStamp > 1 Then
        Return 1
    End If

    Return timeStamp
End Script

' 计算已分段的 t 参数
' @params t 当前全局时间参数
' @return 返回范围为 [0, 1] 的参数
Export Script CUTimeCalcSplitedT(t As Double, Return Double)
    t = t * libCurver_timeStampSplitCnt
    Return t - Int(t)
End Script

' 计算当前所处段号
' @params t 当前全局时间参数
' @return 当前所处段号
Export Script CUTimeCalcSplitIdx(t As Double, Return Long)
    Return Int(t * libCurver_timeStampSplitCnt)
End Script

' 计算已分段 t 参数 (t∈[0, 1])
' @params t 当前全局时间参数
' @params curSplitIdx 当前段号
' @return 若当前 t 不属于当前段号则返回 -1, 否则返回范围为 [0, 1] 的参数
Export Script CUTimeCalcSplitedTByIdx(t As Double, curSplitIdx As Long, Return Double)
    If curSplitIdx <= 0 Or curSplitIdx > libCurver_timeStampSplitCnt - 1 Then
        Return -1
    End If

    t = t * libCurver_timeStampSplitCnt
    If t < curSplitIdx Or t - curSplitIdx > 1.000000001 Then
        Return -1
    End If

    Return t - curSplitIdx
End Script

' 计算已分段的 t 参数
' @params timeStamp 当前时间戳
' @return 返回范围为 [0, 1] 的参数
Export Script CUTimeCalcSplitedTByStamp(timeStamp As Double, Return Double)
    timeStamp = CUTimeCalcT(timeStamp)
    timeStamp = timeStamp * libCurver_timeStampSplitCnt
    Return timeStamp - Int(timeStamp)
End Script

' 计算当前所处段号
' @params timeStamp 当前时间戳
' @return 当前所处段号
Export Script CUTimeCalcSplitIdxByStamp(timeStamp As Double, Return Long)
    timeStamp = CUTimeCalcT(timeStamp)
    Return Int(timeStamp * libCurver_timeStampSplitCnt)
End Script

' 计算已分段 t 参数 (t∈[0, 1])
' @params timeStamp 当前时间戳
' @params curSplitIdx 当前段号
' @return 若当前 t 不属于当前段号则返回 -1, 否则返回范围为 [0, 1] 的参数
Export Script CUTimeCalcSplitedTByStampAndIdx(timeStamp As Double, curSplitIdx As Long, Return Double)
    timeStamp = CUTimeCalcT(timeStamp)
    If curSplitIdx <= 0 Or curSplitIdx > libCurver_timeStampSplitCnt - 1 Then
        Return -1
    End If

    timeStamp = timeStamp * libCurver_timeStampSplitCnt
    If timeStamp < curSplitIdx Or timeStamp - curSplitIdx > 1.000000001 Then
        Return -1
    End If

    Return timeStamp - curSplitIdx
End Script

' ----------------------- 时间工具集 END
' ------------------------------------

' ------------------------------------
' ----------------------- Ease t BEGIN

Export Script CUEase_InSine(t As Double, Return Double)
    Return -Cos(t * 1.5707964) + 1
End Script

Export Script CUEase_OutSine(t As Double, Return Double)
    Return Sin(t * 1.5707964)
End Script

Export Script CUEase_InOutSine(t As Double, Return Double)
    Return -0.5 * (Cos(3.14159265357 * t) - 1)
End Script

Export Script CUEase_InQuad(t As Double, Return Double)
    Return t * t
End Script

Export Script CUEase_OutQuad(t As Double, Return Double)
    Return -t * (t - 2)
End Script

Export Script CUEase_InOutQuad(t As Double, Return Double)
    t = t * 2
    If t < 1 Then
        Return 0.5 * t * t
    End If
    t = t - 1
    Return -0.5 * (t * (t - 2) - 1)
End Script

Export Script CUEase_InCubic(t As Double, Return Double)
    Return t * t * t
End Script

Export Script CUEase_OutCubic(t As Double, Return Double)
    t = t - 1
    Return t * t * t + 1
End Script

Export Script CUEase_InOutCubic(t As Double, Return Double)
    t = t * 2
    If t < 1 Then
        Return 0.5 * t * t * t
    End If
    t = t - 2
    Return 0.5 * (t * t * t + 2)
End Script

Export Script CUEase_InQuart(t As Double, Return Double)
    Return t * t * t * t
End Script

Export Script CUEase_OutQuart(t As Double, Return Double)
    t = t - 1
    t = t * t
    Return -t * t + 1
End Script

Export Script CUEase_InOutQuart(t As Double, Return Double)
    t = t * 2
    If t < 1 Then
        Return 0.5 * t * t * t * t
    End If
    t = t - 2
    Return -0.5 * (t * t * t * t - 2)
End Script

Export Script CUEase_InExpo(t As Double, Return Double)
    If Abs(t) < 0.000000001 Then
        Return 0
    End If
    Return 2 ^ (10 * (t - 1))
End Script

Export Script CUEase_OutExpo(t As Double, Return Double)
    If Abs(t - 1) < 0.000000001 Then
        Return 0
    End If
    Return - (2 ^ (-10 * t)) + 1
End Script

Export Script CUEase_InOutExpo(t As Double, Return Double)
    If Abs(t) < 0.000000001 Then
        Return 0
    End If
    If Abs(t - 1) <0.000000001 Then
        Return 1
    End If

    t = t * 2
    If t < 1 Then
        Return 0.5 * (2 ^ (10 * t - 1))
    End If
    Return 0.5 * (2 ^ (-10 * (t - 1) + 2))
End Script

Export Script CUEase_InCirc(t As Double, Return Double)
    Return - Sqr(1 - t * t) - 1
End Script

Export Script CUEase_OutCirc(t As Double, Return Double)
    t = t - 1
    Return Sqr(1 - t * t)
End Script

Export Script CUEase_InOutCirc(t As Double, Return Double)
    t = t * 2
    If t < 1 Then
        Return - 0.5 * (Sqr(1 - t * t) - 1)
    End If
    t = t - 2
    Return 0.5 * (Sqr(1 - t * t) + 1)
End Script

' ----------------------- Ease t END
' ------------------------------------

' ------------------------------------
' ----------------------- Math END

Export Script CUMath_Max(a As Double, b As Double, Return Double)
    If a >= b Then
        Return a
    Else
        Return b
    End If
End Script

Export Script CUMath_Min(a As Double, b As Double, Return Double)
    If a <= b Then
        Return a
    Else
        Return b
    End If
End Script

Export Script CUMath_Clamp01(a As Double, Return Double)
    If a < 0 Then
        Return 0
    ElseIf a > 1 Then
        Return 1
    Else
        Return a
    End If
End Script

Export Script CUMath_Lerp(a As Double, b As Double, t As Double, Return Double)
    Return a * (1 - t) + t * b
End Script

' ----------------------- Math END
' ------------------------------------