#### Rechnungen - FBA (Fullfillment by Amazon)

function sendmail ($to, $content)
{
	$mail = New-Object System.Net.Mail.MailMessage
	Write-Host "Sending mail to $to"
	$mail.From = "FBA@warensortiment.de"
	$mail.To.Add($to)
	$mail.Subject = "Amazon Fulfillment Report"
	$mail.Body = $content
	$smtp = New-Object System.Net.Mail.SmtpClient("mail.warensortiment.de");
	$smtp.Credentials = New-Object System.Net.NetworkCredential("gsc@warensortiment.de", "01716100817");
	$smtp.Send($mail);
	$mail = $null
	$smtp = $null
} 

# Customer:
# !GENERATION_DATE!
# !GENERATOR_INFO!
# !CUSTOMER_ID!
# !EXTERNAL_ID!
# !COUNTRY!
# !FIRST_NAME!
# !LAST_NAME!
# !EMAIL_ADDRESS!
# !TELEFON_NUMMER!
# !STRASSE_1!
# !STRASSE_2!
# !POSTLEITZAHL!
# !STADT!
# !STEUERDEFINITION! (Inland usw.)

function get_customer_id ()
{
	[xml]$saved_variables = (get-content D:\import\amazon\xml\saved_variables.xml) ### auslesen der nächsten kd-nr
	[int]$customer_id = $saved_variables.Customers.next_customer_id
	$next_customer_id = $customer_id +1
	$saved_variables.Customers.next_customer_id = $next_customer_id.ToString()
	$saved_variables.Save("D:\import\amazon\xml\saved_variables.xml")

	return $customer_id
}

function get_steuer_def ($kuerzel)
{
	switch($kuerzel)
	{
		"SK" {return "EU-ohneUstID"}
		"SI" {return "EU-ohneUstID"}
		"SE" {return "EU-ohneUstID"}
		"PT" {return "EU-ohneUstID"}
		"RO" {return "EU-ohneUstID"}
		"PL" {return "EU-ohneUstID"}
		"NL" {return "EU-ohneUstID"}
		"MT" {return "EU-ohneUstID"}
		"LV" {return "EU-ohneUstID"}
		"LU" {return "EU-ohneUstID"}
		"LT" {return "EU-ohneUstID"}
		"IE" {return "EU-ohneUstID"}
		"HU" {return "EU-ohneUstID"}
		"GR" {return "EU-ohneUstID"}
		"GB" {return "EU-ohneUstID"}
		"FR" {return "EU-ohneUstID"}
		"FI" {return "EU-ohneUstID"}
		"ES" {return "EU-ohneUstID"}
		"EE" {return "EU-ohneUstID"}
		"DK" {return "EU-ohneUstID"}
		"CZ" {return "EU-ohneUstID"}
		"CY" {return "EU-ohneUstID"}
		"BG" {return "EU-ohneUstID"}
		"IT" {return "EU-ohneUstID"}
		"DE" {return "Inland"}
		"BE" {return "EU-ohneUstID"}
		"AT" {return "EU-ohneUstID"}
		default {return "Drittland"}
	}
}


#####################################################
#####################################################
################    START   #########################
#####################################################
#####################################################
$customer_template = "D:\import\amazon\xml\FBA_customer.xml"
$order_template = "D:\import\amazon\xml\FBA_order.xml"
$folder = "D:\import\amazon-fullfilment\"
$location_customer = "D:\import\amazon\fullfilment\customer\"
$location_order = "D:\import\amazon\fullfilment\order"
$dateinamen = get-ChildItem $folder\*.* -name -include *.txt

foreach ($datei in $dateinamen)
{
	if ($datei)
	{
		echo $folder\$datei
		$data = Import-Csv $folder\$datei -Delimiter "	" #-Header "Kunden Auftragsnr. (max. 20 Zeichen)","Abhol Anrede (max. 6 Zeichen)","*Abhol Name (max. 50 Zeichen)","Abhol Zusatz (max. 50 Zeichen)","*Abhol Strasse (max. 50 Zeichen)","*Abhol HNr.","*Abhol PLZ","*Abhol Ort (max. 50 Zeichen)","Abhol Bemerkung (max. 50 Zeichen)","*Ausfuehrungsdatum","*Sperrgut (J, j, N, n, L oder l )","Sperrgut Text (max. 20 Zeichen, Angabe bei Sperrgutsendungen)","Identcode (genau 12 Zeichen)","Leitcode","Auftraggeber Anrede","*Auftraggeber Name (max. 50 Zeichen)","Auftraggeber Zusatz (max. 50 Zeichen)","*Auftraggeber Strasse (max. 50 Zeichen)","*Auftraggeber HNr.","*Auftraggeber PLZ","*Auftraggeber Ort (max. 50 Zeichen)","Auftraggeber Bemerkung (max. 50 Zeichen)","Empfänger Anrede","*Empfänger Name (max. 50 Zeichen)","Empfänger Zusatz (max. 50 Zeichen)","*Empfänger Strasse (max. 50 Zeichen)","*Empfänger HNr.","*Empfänger PLZ","*Empfänger Ort (max. 50 Zeichen)","Empfänger Bemerkung (max. 50 Zeichen)"



		for($i = 0; $i -lt $data.Length; $i += 1)
		{   
            ########################
            ####### CUSTOMER #######
            ########################
            
			$order_id = $data[$i]."amazon-order-id"
			$generation_date = get-date -uFormat “%Y%m%d%H%M%S”
			$generator_info = "PCE FBA Importer Beta 1"
			$customer_id = get_customer_id
			$external_id = "Amazon FBA $order_id"
			$country = $data[$i]."bill-country"
			$name = $data[$i]."recipient-name"
			$name = $name.Split(" ")
			
			
			if($name[2])
			{
			$last_name = $name[2]
			$first_name = $name[0]
			$first_name += ' '
			$first_name += $name[1]
			}
			else
			{
			$last_name = $name[1]
			$first_name = $name[0]
			}
			$email_address = $data[$i]."buyer-email"
			$telefon_nummer = $data[$i]."ship-phone-number"
			$Strasse_1 = $data[$i]."bill-address-1"
			$Strasse_2 = $data[$i]."bill-address-2"
			$Strasse_2 += "     "
			$Strasse_2 += $data[$i]."bill-address-3"
			$postleitzahl = $data[$i]."bill-postal-code"
			$stadt = $data[$i]."bill-city"
			$steuerdefinition = get_steuer_def $country

			$content = Get-Content $customer_template
			$content = $content -replace "!GENERATION_DATE!", $generation_date
			$content = $content -replace "!GENERATOR_INFO!", $generator_info
			$content = $content -replace "!CUSTOMER_ID!", $customer_id
			$content = $content -replace "!EXTERNAL_ID!", $external_id
			$content = $content -replace "!COUNTRY!", $country
			$content = $content -replace "!FIRST_NAME!", $first_name
			$content = $content -replace "!LAST_NAME!", $last_name
			$content = $content -replace "!EMAIL_ADDRESS!", $email_address
			$content = $content -replace "!TELEFON_NUMMER!", $telefon_nummer
			$content = $content -replace "!STRASSE_1!", $strasse_1
			$content = $content -replace "!STRASSE_2!", $strasse_2
			$content = $content -replace "!POSTLEITZAHL!", $postleitzahl
			$content = $content -replace "!STADT!", $stadt
			$content = $content -replace "!STEUERDEFINITION!", $steuerdefinition
            
			$content = $content -replace "&", "u."
			
			$content | out-file -Encoding "UTF8" -filepath $location_customer$order_id"_customer".xml
            
            #######################
            #######  ORDER  #######
            #######################
            
            $purchase_date = $data[$i]."purchase-date"
            
            
            $ship_street_1 = $data[$i]."ship-address-1"
			$ship_street_2 = $data[$i]."ship-address-2"
			$ship_street_2 += "     "
			$ship_street_2 += $data[$i]."ship-address-3"
            $ship_name = $data[$i]."recipient-name"
            $ship_city = $data[$i]."ship-city"
            $ship_state = $data[$i]."ship-state"
            $ship_zip = $data[$i]."ship-postal-code"
            $ship_country = $data[$i]."ship-country"
            
            $bill_street_1 = $data[$i]."bill-address-1"
			$bill_street_2 = $data[$i]."bill-address-2"
			$bill_street_2 += "     "
			$bill_street_2 += $data[$i]."bill-address-3"
            $bill_city = $data[$i]."bill-city"
            $bill_state = $data[$i]."bill-state"
            $bill_zip = $data[$i]."bill-postal-code"
            $bill_country = $data[$i]."bill-country"
            
            $versandkosten = 0
            $versandkosten += $data[$i]."shipping-price"
            $versandkosten += $data[$i]."gift-wrap-price"
			$versandkosten = (($versandkosten)/119)*100
			
			$price =(($data[$i]."item-price")/119)*100
			$itemcount = $data[$i]."quantity-shipped"
			$total_price = $price*$itemcount
			
			$item = $itemcount = $data[$i]."sku"
            
            $positionsmenge = $data[$i]."quantity-shipped"
            
            $content = Get-Content $order_template
            
            $content = $content -replace "!GENERATION_DATE!", $generation_date
			$content = $content -replace "!GENERATOR_INFO!", $generator_info
            $content = $content -replace "!EXTERNAL_ID!", $external_id
            $content = $content -replace "!ORDER_DATE!", $purchase_date
            $content = $content -replace "!FIRST_NAME!", $first_name
			$content = $content -replace "!LAST_NAME!", $last_name
            $content = $content -replace "!STRASSE_1!", $strasse_1
			$content = $content -replace "!STRASSE_2!", $strasse_2
            $content = $content -replace "!POSTLEITZAHL!", $postleitzahl
            $content = $content -replace "!STADT!", $stadt
            $content = $content -replace "!COUNTRY!", $country
            $content = $content -replace "!TELEFON_NUMMER!", $telefon_nummer
            $content = $content -replace "!CUSTOMER_ID!", $customer_id
			$content = $content -replace "!EMAIL_ADDRESS!", $email_address
            
            $content = $content -replace "!SHIP_STREET_1!", $ship_street_1
			$content = $content -replace "!SHIP_STREET_2!", $ship_street_2
            $content = $content -replace "!SHIP_NAME!", $ship_name
            $content = $content -replace "!SHIP_ZIP!", $ship_zip
            $content = $content -replace "!SHIP_CITY!", $ship_city
            $content = $content -replace "!SHIP_COUNTRY!", $ship_country  
            
            $content = $content -replace "!BILL_STREET_1!", $bill_street_1
			$content = $content -replace "!BILL_STREET_2!", $bill_street_2
            $content = $content -replace "!BILL_ZIP!", $bill_zip
            $content = $content -replace "!BILL_CITY!", $bill_city
            $content = $content -replace "!BILL_COUNTRY!", $bill_country 
			
            $content = $content -replace "!VERSANDKOSTEN!", $versandkosten
            
			$content = $content -replace "!ITEM!", $item
			$content = $content -replace "!PRICE!", $price
			$content = $content -replace "!ITEMCOUNT!", $itemcount
			$content = $content -replace "!TOTAL_PRICE!", $total_price
			$content = $content -replace "!POSITIONSMENGE!", $positionsmenge
			
            $content = $content -replace "&", "u."
			
            $content | out-file -Encoding "UTF8" -filepath $location_order$order_id"_order_"$i.xml
		}
	}
    
move-item $folder\$datei $folder\old\

echo "sleeping..."
    sleep 500

}

$dateinamen = get-ChildItem D:\import\amazon\fullfilment\*.* -name -include *.xml

foreach ($datei in $dateinamen)
	{
    if ($datei)
        {
			move-item D:\import\amazon\fullfilment\$datei D:\import\amazon\fullfilment\order\
        }
}