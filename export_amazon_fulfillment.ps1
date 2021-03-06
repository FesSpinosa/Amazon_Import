############################################
# Amazon XML export for powershell v2.0    #
# - Amazon -                               #
# Author: Georg Schröjahr                  #
# PCE Deutschland GmbH, Im Langel 4        #
# 59872 Meschede                           #
# 01/10                                    #
# All rights reserved.                     #
############################################
$time = get-date -uFormat “%d.%m.%Y %H:%M:%S”
$queue = get-content D:\import\amazon\xml\sending_queue.txt 
$dateinamen = get-ChildItem D:\dhl\export\standard\archive\*.* -name -include *.txt

foreach ($datei in $dateinamen)
	{
    if ($datei)
        {
            $csv = Import-Csv D:\dhl\export\standard\archive\$datei -Delimiter ";" -Header "1","PNR","3","4","5","6","7","8","9","10","11","12","13","LS","15","16","17","18","19","20","21","22","KDNR","24","25","26","27","28","29","30","31","32","33","34","35","36","37","38","39","40","41","42","43","44","45","46","47","48","49","50","51","52","53"
            $KDNR = $csv.KDNR

            $queue = import-csv D:\import\amazon\xml\sending_queue.txt
            $separated = $queue  | Where-Object {$_.Kunde -eq $KDNR}

            if ($separated)
                {
                    [xml]$versandbestaetigung = (get-content D:\import\amazon\xml\versandbestaetigung.xml)
                    $temp = $separated.Data
                    $versandbestaetigung.AmazonEnvelope.Message.OrderFulfillment.MerchantOrderID = $separated.Data
                    $versandbestaetigung.AmazonEnvelope.Message.OrderFulfillment.MerchantFulfillmentID = $separated.Data
                    $versandbestaetigung.AmazonEnvelope.Message.OrderFulfillment.FulfillmentDate = "$time"
                    $versandbestaetigung.AmazonEnvelope.Message.OrderFulfillment.Item.MerchantFulfillmentItemID = $separated.Data
                    $versandbestaetigung.AmazonEnvelope.Message.OrderFulfillment.Item.MerchantOrderItemID = ($separated.Data.GetHashCode()*-1).toString()
                    $versandbestaetigung.AmazonEnvelope.Message.OrderFulfillment.FulfillmentData.ShipperTrackingNumber = $csv.PNR
                    $versandbestaetigung.Save("D:\import\amazon\amazon-amtu\DocumentTransport\production\outgoing\order-fulfillment-$temp.xml")
                    $queue -notmatch $KDNR | export-csv D:\import\amazon\xml\sending_queue.txt
                    echo "removed $KDNR from queue file"
                    echo "created D:\import\amazon\amazon-amtu\DocumentTransport\production\order-fulfillment-$temp.xml"
                                                            
                }
           
           
            move-item D:\dhl\export\standard\archive\$datei D:\dhl\export\standard\processed\   
        }
    }