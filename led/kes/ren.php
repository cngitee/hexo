<?php
echo  "Session：".session_id();//获取唯一session值
echo "<br>";
if(getenv('HTTP_CLIENT_IP')) {
	$onlineip = getenv('HTTP_CLIENT_IP');
} elseif(getenv('HTTP_X_FORWARDED_FOR')) {
    $onlineip = getenv('HTTP_X_FORWARDED_FOR');
} elseif(getenv('REMOTE_ADDR')) {
    $onlineip = getenv('REMOTE_ADDR');
} else {
    $onlineip = $HTTP_SERVER_VARS['REMOTE_ADDR'];
}
//echo "你的IP地址：" . $onlineip;
$ip4 = sprintf('%u',ip2long($onlineip)); // 转int数字结果
echo "真实IP：".$ip4;//输出真实IP结果
echo "<br>";

date_default_timezone_set(PRC);     //将date函数默认时间设置中国区时间
$currenttime=date("Y-m-d H:i:s");   //给变量赋值，调用date函数，格式为 年-月-日 时:分:秒
$date = new DateTime(''.$currenttime);//变量取实时时间
$timestamp = $date->getTimestamp();//转时间戳
echo "时间戳："."$timestamp";//输出时间戳