function CommmaTxt ($line, $i){
    $tline = ""
    $tmp = $line.Split(",")
    $tmp[0] = $i
    foreach($ttmp in $tmp){
        if ($tline -eq ""){
            $tline = $tline + $ttmp
        }
        else{
            $tline = $tline + "," + $ttmp
        }
    }
    return $tline
}

function OneTeam ($FR){
    $l = 0
    while (($line = $FR.ReadLine()) -ne $null)
    {
        $tline = ""
        if($l -ne 0){
            $tmp = $line.Split(",")
            $tmp[0] = 0
            foreach($ttmp in $tmp){
                if ($tline -eq ""){
                    $tline = $tline + $ttmp
                }
                else{
                    $tline = $tline + "," + $ttmp
                 }
             }
             $Global:CIOA += $tline
        }
        $l++
    }
}


$path = Split-Path -Parent $MyInvocation.MyCommand.Path

#$Args[0] 第1引数はハンマーモデル結果出力パス
$path = $Args[0]

Set-Location $path
$CI0FN = $path + "\CI_0.csv"
$CI1FN = $path + "\CI_1.csv"
$CI2FN = $path + "\CI_2.csv"
$CI3FN = $path + "\CI_3.csv"

$CI0R = New-Object System.IO.StreamReader($CI0FN, [System.Text.Encoding]::GetEncoding("sjis"))
$CI1R = New-Object System.IO.StreamReader($CI1FN, [System.Text.Encoding]::GetEncoding("sjis"))
$CI2R = New-Object System.IO.StreamReader($CI2FN, [System.Text.Encoding]::GetEncoding("sjis"))
$CI3R = New-Object System.IO.StreamReader($CI3FN, [System.Text.Encoding]::GetEncoding("sjis"))

# 各クラスタの人数を配列に格納
$CIXF = @()
$CIXF += (Get-Content -Path $CI0FN).Length -1 #header行除く
$CIXF += (Get-Content -Path $CI1FN).Length -1 #header行除く
$CIXF += (Get-Content -Path $CI2FN).Length -1 #header行除く
$CIXF += (Get-Content -Path $CI3FN).Length -1 #header行除く

# Header取得
$HDR = Get-Content -Path $CI0FN -totalcount 1

# Write用格納配列
$Global:CIOA = @()
$Global:CIOA += $HDR

# 総人数
$NoP=0
for ($i=0; $i -lt $CIXF.Count; $i++){
    $NoP = $NoP + $CIXF[$i]
}


if($NoP -lt 8){
    $GNM = 1
}
elseif($NoP -lt 13){
    $GNM = 2
}
elseif($NoP -lt 16){
    $GNM = 3
}
elseif($NoP -lt 20){
    $GNM = 4
}
else{
     $GNM = $NoP / 5
     if($NoP % 5 -eq 0){

     }
     if($NoP % 5 -eq 1){

     }
     if($NoP % 5 -eq 2){

     }
     if($NoP % 5 -eq 3){

     }
     if($NoP % 5 -eq 4){

     }
}



if($NoP -lt 8){
    OneTeam $CI0R $Global:CIOA
    OneTeam $CI1R $Global:CIOA
    OneTeam $CI2R $Global:CIOA
    OneTeam $CI3R $Global:CIOA
}
elseif($NoP -ge 20){
    $s = $NoP / 5
    $t = $NoP % 5
    $ta = @()
    for($i=0; $i -lt $s; $i++){
        # 余りが0の時 all 5
        # 余りが1の時 1個目6 その他 5
        # 余りが2の時 2個目まで6 その他 5
        # 余りが3の時 3個目まで6 その他 5
        # 余りが4の時 4個目まで6 その他 5
        if($i -eq 0){
            if($t -eq 0){
                $ta += 5
            }
            else{
                $ta += 6
            }
        }
        elseif($i -eq 1){
            if($t -gt 1){
                $ta += 6
            }
            else{
                $ta += 5
            }
        }
        elseif($i -eq 2){
            if($t -gt 2){
                $ta += 6
            }
            else{
                $ta += 5
            }
        }
        elseif($i -eq 3){
            if($t -gt 3){
                $ta += 6
            }
            else{
                $ta += 5
            }
        }
        else{
            $ta += 5
        }
    }
    # チーム分け
    $line = $CI0R.ReadLine()
    $line = $CI1R.ReadLine()
    $line = $CI2R.ReadLine()
    $line = $CI3R.ReadLine()

    $CI0L = 1
    $CI1L = 1
    $CI2L = 1
    $CI3L = 1

#    for($i=0; $i -lt $NoP; $i++){
#        if($CIXF[0] -gt 0){
#        }
#    }
    for($i=0; $i -lt $ta.Count; $i++){
        $u=0
        if($CIXF[0] -gt 0){
            $line = $CI0R.ReadLine()
            $Global:CIOA += (CommmaTxt $line $i)
            $u++
            $CIXF[0]--
        }
        if($CIXF[1] -gt 0){
            $line = $CI1R.ReadLine()
            $Global:CIOA += (CommmaTxt $line $i)
            $u++
            $CIXF[1]--
        }
        if($CIXF[2] -gt 0){
            $line = $CI2R.ReadLine()
            $Global:CIOA += (CommmaTxt $line $i)
            $u++
            $CIXF[2]--
        }
        if($CIXF[3] -gt 0){
            $line = $CI3R.ReadLine()
            $Global:CIOA += (CommmaTxt $line $i)
            $u++
            $CIXF[3]--
        }
        while($u -lt $ta[$i]){
            $v=0
            for($w=1; $w -lt 4; $w++){
                if($CIXF[$v] -lt $CIXF[$w]){
                    $v=$w
                } 
            }
            if($v -eq 0){
                if($CIXF[0] -gt 0){
                    $line = $CI0R.ReadLine()
                    $Global:CIOA += (CommmaTxt $line $i)
                    $CIXF[0]--
                }
            }
            elseif($v -eq 1){
                $line = $CI1R.ReadLine()
                $Global:CIOA += (CommmaTxt $line $i)
                $CIXF[1]--
            }
            elseif($v -eq 2){
                $line = $CI2R.ReadLine()
                $Global:CIOA += (CommmaTxt $line $i)
                $CIXF[2]--
            }
            elseif($v -eq 3){
                $line = $CI3R.ReadLine()
                $Global:CIOA += (CommmaTxt $line $i)
                $CIXF[3]--
            }
            $u++
        }
    }
}


#$Args[1] 第2引数は出力パス

$Global:CIOA | Out-File -FilePath $Args[1]


$CI0R.Close()
$CI1R.Close()
$CI2R.Close()
$CI3R.Close()
