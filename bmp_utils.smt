' ===================================================
' ================================ bitmap 工具类 v0.1
' ================================== xiaodou 20240101
' ===================================================

' ===================================== 临时变量

' ------------------------ bitmap 定义变量
Dim libbmp_angle As Double = 0
Dim libbmp_isVisible As Byte = 1
Dim libbmp_isUseScreenCoord As Byte = 0

Dim libbmp_sw As Long       = 1
Dim libbmp_sh As Long       = 1
Dim libbmp_sx As Long       = 0
Dim libbmp_sy As Long       = 0
Dim libbmp_srcId As Long    = 1

Dim libbmp_destX As Long    = 0
Dim libbmp_destY As Long    = 0

Dim libbmp_scaleW As Double  = 1
Dim libbmp_scaleH As Double  = 1

Dim libbmp_anchorX As Double = 0
Dim libbmp_anchorY As Double = 0

Dim libbmp_animW As Long = 0
Dim libbmp_animH As Long = 0
Dim libbmp_animStartX As Long = 0
Dim libbmp_animStartY As Long = 0

Dim libbmp_animCnt As Long = 1
Dim libbmp_animCeilX As Long = 0
Dim libbmp_animIntervalX As Long = 0
Dim libbmp_animIntervalY As Long = 0
Dim libbmp_animLoopFrameCnt As Long = 0

Dim libbmp_col_r1 As Integer = 0
Dim libbmp_col_g1 As Integer = 0
Dim libbmp_col_b1 As Integer = 0
Dim libbmp_col_a1 As Integer = 0

Dim libbmp_col_r2 As Integer = 0
Dim libbmp_col_g2 As Integer = 0
Dim libbmp_col_b2 As Integer = 0
Dim libbmp_col_a2 As Integer = 0

Dim libbmp_offsetAnimW As Long = 0
Dim libbmp_offsetAnimH As Long = 0
Dim libbmp_offsetAnimStartX As Long = 0
Dim libbmp_offsetAnimStartY As Long = 0

' ------------------------ 临时变量
Dim libbmp_temp_i As Double = 0
Dim libbmp_temp_j As Double = 0
Dim libbmp_temp_k As Double = 0
Dim libbmp_temp_l As Double = 0
Dim libbmp_temp_m As Double = 0
Dim libbmp_temp_n As Double = 0
Dim libbmp_temp_o As Double = 0
Dim libbmp_temp_p As Double = 0
Dim libbmp_temp_q As Double = 0
Dim libbmp_temp_r As Double = 0
Dim libbmp_temp_s As Double = 0
Dim libbmp_temp_t As Double = 0

' ===================================== 创建 bmp

' 设置 srcId
Export Script BmpNewStoreSrcId(srcId As Long)
    libbmp_srcId = srcId
End Script

' 设置 src 采样偏移和大小
Export Script BmpNewStoreSrcOffset(x As Long, y As Long, w As Long, h As Long)
    libbmp_sx = x
    libbmp_sy = y
    libbmp_sh = h
    libbmp_sw = w
End Script

' 设置 src 信息
Export Script BmpNewStoreSrc(srcId As Long, x As Long, y As Long, w As Long, h As Long)
    Call BmpNewStoreSrcId(srcId)
    Call BmpNewStoreSrcOffset(x, y, w, h)
End Script

' 设置 bmp 缩放
Export Script BmpNewStoreScale(w As Long, h As Long)
    libbmp_scaleW = w
    libbmp_scaleH = h
End Script

' 设置 bmp 是否使用屏幕坐标
Export Script BmpNewStoreIsUseScreenCoords(isUse As Byte)
    libbmp_isUseScreenCoord = isUse
End Script

' 设置 bmp 位置
Export Script BmpNewStorePos(x As Long, y As Long)
    libbmp_destX = x
    libbmp_destY = y
End Script

' 设置 bmp 角度
Export Script BmpNewStoreAngle(angle As Double)
    libbmp_angle = angle
End Script

' 创建 bmp
Export Script BmpNew(id As Long)
    Call BMPCreate(id, libbmp_srcId, libbmp_isUseScreenCoord, libbmp_isVisible, libbmp_sx, libbmp_sy, libbmp_sw, libbmp_sh, libbmp_destX, libbmp_destY, libbmp_scaleW, libbmp_scaleH, 0, 0, libbmp_angle, -1)
End Script

' 销毁 bmp
Export Script BmpDel(id As Long)
    Call BErase(2, id)
End Script

' ===================================== 设置 bmp

' 回复旋转
Script BmpInner_RotateReset(id As Long, Return Integer)
    libbmp_temp_m = Bitmap(id).rotatang
    If libbmp_temp_m = 0 Then
        Return -1
    End If
    libbmp_temp_i = Bitmap(id).scrwidth * Bitmap(id).scalex * libbmp_anchorX
    libbmp_temp_j = Bitmap(id).scrheight * Bitmap(id).scaley * libbmp_anchorY
    libbmp_temp_k = libbmp_temp_i * Cos(libbmp_temp_m) + libbmp_temp_j * Sin(libbmp_temp_m)
    libbmp_temp_l = -libbmp_temp_i * Sin(libbmp_temp_m) + libbmp_temp_j * Cos(libbmp_temp_m)

    Bitmap(id).destx += libbmp_temp_k - libbmp_temp_i
    Bitmap(id).desty += libbmp_temp_l - libbmp_temp_j
    Bitmap(id).rotatang = 0
    Return 0
End Script

' 储存相对锚点
Export Script BmpStoreAnchor(x As Double, y As Double)
    libbmp_anchorX = x
    libbmp_anchorY = y
End Script

' 设置位置
Export Script BmpPos(id As Long, x As Long, y As Long)
    If (Abs(libbmp_anchorX) > 0.0001 Or Abs(libbmp_anchorY) > 0.0001) And Bitmap(id).scrwidth > 0 And Bitmap(id).scrheight > 0 Then
        libbmp_temp_o = Bitmap(id).rotatx
        libbmp_temp_p = Bitmap(id).rotaty
        Bitmap(id).rotatx = 0
        Bitmap(id).rotaty = 0

        libbmp_temp_n = Bitmap(id).rotatang
        Call BmpInner_RotateReset(id)
        Bitmap(id).destx = x - Bitmap(id).scrwidth * Bitmap(id).scalex * libbmp_anchorX
        Bitmap(id).desty = y - Bitmap(id).scrheight * Bitmap(id).scaley * libbmp_anchorY
        Call BmpRotate(id, libbmp_temp_n)
        Bitmap(id).rotatx = libbmp_temp_o
        Bitmap(id).rotaty = libbmp_temp_p
    Else
        Bitmap(id).destx = x
        Bitmap(id).desty = y
    End If
End Script

' 设置转角
Export Script BmpRotate(id As Long, angle As Double)
    If (Abs(libbmp_anchorX) > 0.0001 Or Abs(libbmp_anchorY) > 0.0001) And Bitmap(id).scrwidth > 0 And Bitmap(id).scrheight > 0 Then
        libbmp_temp_q = Bitmap(id).rotatx
        libbmp_temp_r = Bitmap(id).rotaty
        Bitmap(id).rotatx = 0
        Bitmap(id).rotaty = 0

        Call BmpInner_RotateReset(id)
        libbmp_temp_k = libbmp_temp_i * Cos(angle) + libbmp_temp_j * Sin(angle)
        libbmp_temp_l = -libbmp_temp_i * Sin(angle) + libbmp_temp_j * Cos(angle)

        Bitmap(id).destx += libbmp_temp_i - libbmp_temp_k
        Bitmap(id).desty += libbmp_temp_j - libbmp_temp_l
        Bitmap(id).rotatang = angle
        Bitmap(id).rotatx = libbmp_temp_q
        Bitmap(id).rotaty = libbmp_temp_r
    Else
        Bitmap(id).rotatang = angle
    End If
End Script

' 设置缩放
Export Script BmpScale(id As Long, scaleX As Double, scaleY As Double)
    If (Abs(libbmp_anchorX) > 0.0001 Or Abs(libbmp_anchorY) > 0.0001) And Bitmap(id).scrwidth > 0 And Bitmap(id).scrheight > 0 Then
        libbmp_temp_s = Bitmap(id).rotatx
        libbmp_temp_t = Bitmap(id).rotaty
        Bitmap(id).rotatx = 0
        Bitmap(id).rotaty = 0

        libbmp_temp_n = Bitmap(id).rotatang
        Call BmpInner_RotateReset(id)
        Bitmap(id).destx -= Bitmap(id).scrwidth * (1 - Bitmap(id).scalex) * libbmp_anchorX
        Bitmap(id).desty -= Bitmap(id).scrheight * (1 - Bitmap(id).scaley) * libbmp_anchorY
        Bitmap(id).scalex = scaleX
        Bitmap(id).scaley = scaleY
        Bitmap(id).destx += Bitmap(id).scrwidth * (1 - Bitmap(id).scalex) * libbmp_anchorX
        Bitmap(id).desty += Bitmap(id).scrheight * (1 - Bitmap(id).scaley) * libbmp_anchorY
        Call BmpRotate(id, libbmp_temp_n)
        Bitmap(id).rotatx = libbmp_temp_s
        Bitmap(id).rotaty = libbmp_temp_t
    Else
        Bitmap(id).scalex = scaleX
        Bitmap(id).scaley = scaleY
    End If
End Script

' ===================================== bmp 动画

' 设置动画序列帧的起始位置
Export Script BmpStoreAnimPos(startX As Long, startY As Long, w As Long, h As Long)
    libbmp_animW = w
    libbmp_animH = h
    libbmp_animStartX = startX
    libbmp_animStartY = startY
End Script

' 设置动画序列帧的数量和横轴数量
Export Script BmpStoreAnimInfo(cnt As Long, cellCntX As Long)
    libbmp_animCnt = cnt
    libbmp_animCeilX = cellCntX
End Script

' 设置动画帧尾循环
Export Script BmpStoreAnimLoop(cnt As Long)
    libbmp_animLoopFrameCnt = cnt
End Script

Export Script BmpStoreAnimInterval(x As Long, y As Long)
    libbmp_animIntervalX = x
    libbmp_animIntervalY = y
End Script

' 设置动画帧
Export Script BmpAnim(id As Long, timeStamp As Long)
    If timeStamp < 0 Then
        timeStamp = 0
    End If

    If libbmp_animLoopFrameCnt > 0 And libbmp_animLoopFrameCnt < libbmp_animCnt And timeStamp >= libbmp_animCnt Then
        timeStamp = (timeStamp mod libbmp_animLoopFrameCnt) + libbmp_animCnt - libbmp_animLoopFrameCnt
    Else
        timeStamp = timeStamp mod libbmp_animCnt
    End If

    libbmp_temp_i = timeStamp mod libbmp_animCeilX
    libbmp_temp_j = timeStamp \ libbmp_animCeilX

    Bitmap(id).scrx = libbmp_animStartX + libbmp_temp_i * (libbmp_animW + libbmp_animIntervalX)
    Bitmap(id).scry = libbmp_animStartY + libbmp_temp_j * (libbmp_animH + libbmp_animIntervalY)
End Script

' ===================================== 颜色

' 设置颜色 1
Export Script BmpStoreCol1(r As Integer, g As Integer, b As Integer, a As Integer)
    libbmp_col_r1 = r
    libbmp_col_g1 = g
    libbmp_col_b1 = b
    libbmp_col_a1 = a
End Script

' 设置颜色 2
Export Script BmpStoreCol2(r As Integer, g As Integer, b As Integer, a As Integer)
    libbmp_col_r2 = r
    libbmp_col_g2 = g
    libbmp_col_b2 = b
    libbmp_col_a2 = a
End Script

' 颜色插值
Export Script BmpLerpCol(t As Double, Return Double)
    Return rgba(((1 - t) * libbmp_col_r1 + t * libbmp_col_r2), ((1 - t) * libbmp_col_g1 + t * libbmp_col_g2), ((1 - t) * libbmp_col_b1 + t * libbmp_col_b2), ((1 - t) * libbmp_col_a1 + t * libbmp_col_a2))
End Script

' 设置 bmp 颜色
Export Script BmpCol(id As Long, r As Integer, g As Integer, b As Integer, a As Integer)
    Bitmap(id).forecolor = rgba(r, g, b, a)
End Script

' 设置 bmp 颜色 1
Export Script BmpCol1(id As Long)
    Bitmap(id).forecolor = rgba(libbmp_col_r1, libbmp_col_g1, libbmp_col_b1, libbmp_col_a1)
End Script

' 设置 bmp 颜色 2
Export Script BmpCol2(id As Long)
    Bitmap(id).forecolor = rgba(libbmp_col_r2, libbmp_col_g2, libbmp_col_b2, libbmp_col_a2)
End Script

' 设置 bmp 颜色插值
Export Script BmpColLerp(id As Long, t As Double)
    Bitmap(id).forecolor = rgba(((1 - t) * libbmp_col_r1 + t * libbmp_col_r2), ((1 - t) * libbmp_col_g1 + t * libbmp_col_g2), ((1 - t) * libbmp_col_b1 + t * libbmp_col_b2), ((1 - t) * libbmp_col_a1 + t * libbmp_col_a2))
End Script

' 设置 bmp alpha 值
Export Script BmpAlpha(id As Long, a As Double)
    Bitmap(id).forecolor_a = a * 255
End Script

' ===================================== 偏移动画

' 设置偏移动画信息
Export Script BmpStoreOffsetAnimInfo(x As Long, y As Long, w As Long, h As Long)
    libbmp_offsetAnimStartX = x
    libbmp_offsetAnimStartY = y
    libbmp_offsetAnimW = w
    libbmp_offsetAnimH = h
End Script

' 设置偏移动画
Export Script BmpOffsetAnim(id As Long, t As Double)
    t = t - Int(t)
    Bitmap(id).scrx = (1 - t) * libbmp_offsetAnimStartX + t * libbmp_offsetAnimW
    Bitmap(id).scry = (1 - t) * libbmp_offsetAnimStartY + t * libbmp_offsetAnimH
End Script