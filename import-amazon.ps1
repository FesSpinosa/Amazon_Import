############################################
# Amazon XML import for powershell v2.0    #
# - Amazon -                               #
# Author: Georg Schröjahr                  #
# PCE Deutschland GmbH, Im Langel 4        #
# 59872 Meschede                           #
# 07/10                                    #
# All rights reserved.                     #
############################################

function sendmail ($to, $content){
    $mail = New-Object System.Net.Mail.MailMessage
    Write-Host "Sending mail to $to"
    $mail.From = "bestellungen@warensortiment.de"
    $mail.To.Add($to)
    $mail.Subject = "$order_count neue Amazon Bestellung(en)"
    $mail.Body = $content
    $smtp = New-Object System.Net.Mail.SmtpClient("85.214.32.39");
    $smtp.Credentials = New-Object System.Net.NetworkCredential("gsc@warensortiment.de", "01716100817");
    $smtp.Send($mail);
    $mail = $null
    $smtp = $null
}    

$time1 = get-date -uFormat “%Y-%m-%d”
$time2 = get-date -uFormat “%d.%m.%Y %H:%M:%S”
$savetime = get-date -uFormat “%d.%m.%Y-%H-%M-%S”

##########################################################################################

$content = ""
$dateinamen = get-ChildItem D:\import\amazon\amazon-amtu\DocumentTransport\production\reports\*.* -name -include *.xml
$process = 0
[System.Reflection.Assembly]::LoadWithPartialName("System.Data.OracleClient")
$connectionString= "Data Source=PCE-ERP1/FET;User Id=FET_USER;Password=FET_USER;Integrated Security=no"
$connection = New-Object System.Data.OracleClient.OracleConnection($connectionString)
$filecount = 0
foreach ($datei in $dateinamen){
    $filecount++
    if ($datei){
        [xml]$xml = (get-content D:\import\amazon\amazon-amtu\DocumentTransport\production\reports\$datei -encoding UTF8) | Foreach-Object {$_ -replace ' xsi:noNamespaceSchemaLocation="amzn-envelope.xsd"', ''}
        $process = 1
        $order_count = $xml.AmazonEnvelope.Message.Length  #Anzahl bestellungen in file bestimmen
        if(!$order_count){
        $order_count = 1
        }
        echo "processing file $datei bestellungen: $order_count"
        for($order_no = 0; $order_no -lt $order_count; $order_no++){
       		if(!$xml.AmazonEnvelope.Message.Length){
            $child = $xml.AmazonEnvelope.Message
            }else{
            $child = $xml.AmazonEnvelope.Message[$order_no]
            }
        	echo "processing... order $filecount"
        	$content += "`n`nBestellung NR.: $filecount `n"
        	[xml]$xml_customer = (get-content D:\import\amazon\xml\blank_customer.xml)
        	[xml]$xml_order = (get-content D:\import\amazon\xml\blank_order.xml)
        	[xml]$bestellbestaetigung = (get-content D:\import\amazon\xml\bestellbestaetigung.xml)
        	$mail = $child.OrderReport.BillingData.BuyerEmailAddress
##### sql
        	$connection.Open()
        	$queryString = "select customer_id from FET_CUSTOMER where customer_proxy_guid = (select GUID from FET_CUSTOMER_PROXY where businesspartner_guid = (select businesspartner_guid from V_BUSINESSPARTNER_MAIL where REGEXP_LIKE(comm_string,'$mail','i')))"
        	$command = new-Object System.Data.OracleClient.OracleCommand($queryString, $connection)
        	$result = $command.ExecuteScalar()   
        	if(!$result){
                $queryString = "select customer_id from FET_CUSTOMER where customer_proxy_guid = (select GUID from FET_CUSTOMER_PROXY where businesspartner_guid = (select businesspartner_guid from FET_BP_EMPLOYEE where guid = (select bp_employee_guid from V_BP_EMPLOYEE_MAIL where REGEXP_LIKE(comm_string,'$mail','i')  and rownum = 1)))"
                $command = new-Object System.Data.OracleClient.OracleCommand($queryString, $connection)
                $result = $command.ExecuteScalar()   
                if(!$result){
                    echo "Kunde nicht vorhanden"
                    $istkundeerstellt = "kundennummer war nicht vorhanden, wurde generiert"
                    [xml]$saved_variables = (get-content D:\import\amazon\xml\saved_variables.xml) ### auslesen der nächsten kd-nr
                    [int]$customer_id = $saved_variables.Customers.next_customer_id
					$next_customer_id = $customer_id +1
                    $next_customer_id
                    $saved_variables.Customers.next_customer_id = $next_customer_id.ToString()
                    $saved_variables.Save("D:\import\amazon\xml\saved_variables.xml")
					$neuer_kunde = $true
                    }else{
                    $istkundeerstellt = "kundennummer war vorhanden, wurde aus der DB genommen"
                    $customer_id = $result
					$neuer_kunde = $false
                    }
                }else{
                $istkundeerstellt = "kundennummer war vorhanden, wurde aus der DB genommen"
                $customer_id = $result
				$neuer_kunde = $false
                }
        	$connection.close()
#### /sql
        	if($neuer_kunde -eq $true){#### übernahme in neuen kunden
			$xml_customer.Customers.ControlInfo.GenerationDate = "$time2"
			$xml_customer.Customers.ControlInfo.GeneratorInfo = "PCE AMAZON XML PARSER BETA 3"
			$xml_customer.Customers.Company.Country = $child.OrderReport.BillingData.Address.CountryCode
			$xml_customer.Customers.Customer.Id.InnerText = "$customer_id"
			$content += "Kundennummer: $customer_id `n"
			$xml_customer.Customers.Customer.Type = "private customer"   #### evtl alternieren
			$xml_customer.Customers.Customer.PrivateCustomerData.Country = $child.OrderReport.BillingData.Address.CountryCode
			$xml_customer.Customers.Customer.PrivateCustomerData.Person.LastName = $child.OrderReport.BillingData.Address.Name
			$content += "Name : "
			$content += $child.OrderReport.BillingData.Address.Name
			$content += "`n"
			$xml_customer.Customers.Customer.PrivateCustomerData.Person.Email.address = $mail
			$xml_customer.Customers.Customer.PrivateCustomerData.Person.Telephone.no = $child.OrderReport.BillingData.Address.PhoneNumber
			$xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.FirstChild.Street = [string]$child.OrderReport.BillingData.Address.AddressFieldOne
			if ($child.OrderReport.BillingData.Address.AddressFieldTwo){
				$xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.FirstChild.Street = $child.OrderReport.BillingData.Address.AddressFieldTwo
				$xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.FirstChild.NameExtension = [string]$child.OrderReport.BillingData.Address.AddressFieldOne
				}
			$xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.FirstChild.Zip = $child.OrderReport.BillingData.Address.PostalCode
			$xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.FirstChild.City = $child.OrderReport.BillingData.Address.City
			$xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.FirstChild.Country = $child.OrderReport.BillingData.Address.CountryCode
			if ($child.OrderReport.BillingData.Address.AddressFieldOne -eq $child.OrderReport.FulfillmentData.Address.AddressFieldOne -AND $child.OrderReport.BillingData.Address.AddressFieldTwo -eq $child.OrderReport.FulfillmentData.Address.AddressFieldTwo){
				$xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.PostalAddress.Roles.Role.Id = "2"
				$xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.PostalAddress.Roles.Role.type = "standard"
				$xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.PostalAddress.Roles.Role.InnerText = "Lieferung"
				$newitem = (@($xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.PostalAddress.Roles.Role)[0]).Clone()
				$newitem.id = "1"
				$newitem.type = "standard"
				$newitem.InnerText = "Auftragsbestätigung"
				$xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.PostalAddress.Roles.AppendChild($newitem) > $null
				$newitem = (@($xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.PostalAddress.Roles.Role)[0]).Clone()
				$newitem.id = "3"
				$newitem.type = "standard"
				$newitem.InnerText = "Rechnung"
				$xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.PostalAddress.Roles.AppendChild($newitem) > $null
				$newitem = (@($xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.PostalAddress.Roles.Role)[0]).Clone()
				$newitem.id = "4"
				$newitem.type = "standard"
				$newitem.InnerText = "Korrespondenz"
				$xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.PostalAddress.Roles.AppendChild($newitem) > $null
				$newitem = (@($xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.PostalAddress.Roles.Role)[0]).Clone()
				$newitem.id = "7"
				$newitem.type = "standard"
				$newitem.InnerText = "Mahnung"
				$xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.PostalAddress.Roles.AppendChild($newitem) > $null          
				}else{
				$xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.PostalAddress.Roles.Role.Id = "3"
				$xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.PostalAddress.Roles.Role.type = "standard"
				$xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.PostalAddress.Roles.Role.InnerText = "Rechnung"
				$newitem = (@($xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.PostalAddress.Roles.Role)[0]).Clone()
				$newitem.id = "1"
				$newitem.type = "standard"
				$newitem.InnerText = "Auftragsbestätigung"
				$xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.PostalAddress.Roles.AppendChild($newitem) > $null
				$newitem = (@($xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.PostalAddress.Roles.Role)[0]).Clone()
				$newitem.id = "4"
				$newitem.type = "standard"
				$newitem.InnerText = "Korrespondenz"
				$xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.PostalAddress.Roles.AppendChild($newitem) > $null
				$newitem = (@($xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.PostalAddress.Roles.Role)[0]).Clone()
				$newitem.id = "7"
				$newitem.type = "standard"
				$newitem.InnerText = "Mahnung"
				$xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.PostalAddress.Roles.AppendChild($newitem) > $null  
				$newitem = (@($xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.PostalAddress)[0]).Clone()
				$newitem.AddressId = "1"
				$newitem.Street = $child.OrderReport.FulfillmentData.Address.AddressFieldOne
				$newitem.Street2 = [string]$child.OrderReport.FulfillmentData.Address.AddressFieldTwo
				$newitem.Zip = $child.OrderReport.FulfillmentData.Address.PostalCode
				$newitem.City = $child.OrderReport.FulfillmentData.Address.City
				$newitem.Country = $child.OrderReport.FulfillmentData.Address.CountryCode
				$newitem.Roles.RemoveChild($newitem.Roles.Role[2]) > $null 
				$newitem.Roles.RemoveChild($newitem.Roles.Role[1]) > $null 
				$newitem.Roles.Role.id = "2"
				$newitem.Roles.Role.type = "standard"
				$newitem.Roles.Role.InnerText = "Lieferung" 
				$xml_customer.Customers.Customer.PrivateCustomerData.PostalAddresses.AppendChild($newitem) > $null              
				}
					
			switch($xml_customer.Customers.Customer.PrivateCustomerData.Country){
				"SK" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"SI" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"SE" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"PT" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"RO" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"PL" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"NL" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"MT" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"LV" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"LU" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"LT" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"IE" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"HU" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"GR" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"GB" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"FR" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"FI" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"ES" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"EE" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"DK" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"CZ" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"CY" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"BG" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"IT" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"DE" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "Inland"}
				"BE" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				"AT" {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "EU-ohneUstID"}
				default {$xml_customer.Customers.Customer.CommonData.TaxDefinition.Name = "Drittland"}
				}
					
			$file_part = $child.OrderReport.AmazonOrderId
			$xml_customer.Save("D:\import\amazon\ausgang\customer\customer-$file_part.xml")
			$xml_customer.Save("D:\import\amazon\ausgang\backup\customer\customer-$file_part.xml")
            }
			$file_part = $child.OrderReport.AmazonOrderId
			
################## Order ##################
            
        	$xml_order.Order.OrderHeader.ControlInfo.GenerationDate = "$time2"
        	$xml_order.Order.OrderHeader.ControlInfo.GeneratorInfo = "PCE AMAZON XML PARSER BETA 1"
        	$xml_order.Order.OrderHeader.OrderId = "Amazon " + $child.OrderReport.AmazonOrderId
#### bestellbestätigung
	        $bestellbestaetigung.AmazonEnvelope.Message.OrderAcknowledgement.AmazonOrderID = $child.OrderReport.AmazonOrderId
    	    $bestellbestaetigung.AmazonEnvelope.Message.OrderAcknowledgement.MerchantOrderID = "Amazon " + $child.OrderReport.AmazonOrderId
        	$temp = ($bestellbestaetigung.AmazonEnvelope.Message.OrderAcknowledgement.MerchantOrderID.GetHashCode()*-1)
        	$bestellbestaetigung.AmazonEnvelope.Message.OrderAcknowledgement.Item.MerchantOrderItemID = $temp.ToString()
        	$merchant_id = $bestellbestaetigung.AmazonEnvelope.Message.OrderAcknowledgement.MerchantOrderID
        	Add-Content D:\import\amazon\xml\sending_queue.txt "$customer_id,$merchant_id"
#### /bestellbestätigung
	       	$xml_order.Order.OrderHeader.OrderDate = $child.OrderReport.OrderDate
    	    $xml_order.Order.OrderHeader.OrderParties.BuyerParty.PartyId.InnerText = $customer_id
        	$xml_order.Order.OrderHeader.OrderParties.BuyerParty.Address.Name = $child.OrderReport.BillingData.Address.Name
        	$xml_order.Order.OrderHeader.OrderParties.BuyerParty.Address.Street = $child.OrderReport.BillingData.Address.AddressFieldOne
			if ($child.OrderReport.BillingData.Address.AddressFieldTwo){
				$xml_order.Order.OrderHeader.OrderParties.BuyerParty.Address.Street = $child.OrderReport.BillingData.Address.AddressFieldTwo
				$xml_order.Order.OrderHeader.OrderParties.BuyerParty.Address.NameExtension = [string]$child.OrderReport.BillingData.Address.AddressFieldOne
			}
        	$xml_order.Order.OrderHeader.OrderParties.BuyerParty.Address.Zip =  $child.OrderReport.BillingData.Address.PostalCode
        	$xml_order.Order.OrderHeader.OrderParties.BuyerParty.Address.City = $child.OrderReport.BillingData.Address.City
        	$xml_order.Order.OrderHeader.OrderParties.BuyerParty.Address.Country = $child.OrderReport.BillingData.Address.CountryCode.ToUpper()
        	$xml_order.Order.OrderHeader.OrderParties.BuyerParty.Contact.ContactName = $child.OrderReport.BillingData.Address.Name
			$xml_order.Order.OrderHeader.OrderParties.BuyerParty.Contact.Phone.innertext = $child.OrderReport.BillingData.BuyerPhoneNumber.tostring() 
			$xml_order.Order.OrderHeader.OrderParties.BuyerParty.Contact.EMail = $child.OrderReport.BillingData.BuyerEmailAddress
### Lieferadresse
	        $xml_order.Order.OrderHeader.OrderParties.DeliveryParty.PartyId.InnerText = $customer_id
    	    $xml_order.Order.OrderHeader.OrderParties.DeliveryParty.Address.Name = $child.OrderReport.FulfillmentData.Address.Name
        	$xml_order.Order.OrderHeader.OrderParties.DeliveryParty.Address.Street = $child.OrderReport.FulfillmentData.Address.AddressFieldOne
			if ($child.OrderReport.FulfillmentData.Address.AddressFieldTwo){
				$xml_order.Order.OrderHeader.OrderParties.DeliveryParty.Address.Street = $child.OrderReport.FulfillmentData.Address.AddressFieldTwo
				$xml_order.Order.OrderHeader.OrderParties.DeliveryParty.Address.NameExtension = [string]$child.OrderReport.FulfillmentData.Address.AddressFieldOne
			}
			$xml_order.Order.OrderHeader.OrderParties.DeliveryParty.Address.Zip =  $child.OrderReport.FulfillmentData.Address.PostalCode
			$xml_order.Order.OrderHeader.OrderParties.DeliveryParty.Address.City = $child.OrderReport.FulfillmentData.Address.City
			$xml_order.Order.OrderHeader.OrderParties.DeliveryParty.Address.Country = $child.OrderReport.FulfillmentData.Address.CountryCode.ToUpper()
### /Lieferadresse
### Rechnungsadresse
			$xml_order.Order.OrderHeader.OrderParties.InvoiceParty.PartyId.InnerText = $customer_id
			$xml_order.Order.OrderHeader.OrderParties.InvoiceParty.Address.Name = $child.OrderReport.BillingData.Address.Name
			$xml_order.Order.OrderHeader.OrderParties.InvoiceParty.Address.Street = $child.OrderReport.BillingData.Address.AddressFieldOne
			if ($child.OrderReport.BillingData.Address.AddressFieldTwo){
				$xml_order.Order.OrderHeader.OrderParties.InvoiceParty.Address.Street = $child.OrderReport.BillingData.Address.AddressFieldTwo
				$xml_order.Order.OrderHeader.OrderParties.InvoiceParty.Address.NameExtension = [string]$child.OrderReport.BillingData.Address.AddressFieldOne
			}
			$xml_order.Order.OrderHeader.OrderParties.InvoiceParty.Address.Zip =  $child.OrderReport.BillingData.Address.PostalCode
			$xml_order.Order.OrderHeader.OrderParties.InvoiceParty.Address.City = $child.OrderReport.BillingData.Address.City
			$xml_order.Order.OrderHeader.OrderParties.InvoiceParty.Address.Country = $child.OrderReport.BillingData.Address.CountryCode.ToUpper()
### /Rechnungsadresse
### Bestellung
			$xml_order.Order.OrderHeader.TermsOfPayment.PaymentType.InnerText = "Amazon"
			$xml_order.Order.OrderHeader.Transport.IncoTerm = "EXW"
			if($child.OrderReport.FulfillmentData.Address.CountryCode.ToLower() -eq "de"){
				$xml_order.Order.OrderHeader.Transport.Name = "Amazon"
				$xml_order.Order.OrderHeader.Transport.Description = "Internetbestellung Amazon"
			}else{
				$xml_order.Order.OrderHeader.Transport.Name = "Amazon Europa"
				$xml_order.Order.OrderHeader.Transport.Description = "Internetbestellung Amazon - europäisches Ausland"
			}
			$newitem = (@($xml_order.order.itemlist.item)[0]).Clone()
			$total = 0
			$itemnum = 0
			$priceamount = 0
			$id = 1
			$single_item_test = $child.OrderReport.Item[0]
			if (!$single_item_test){
				$itemcount = 1
			}else{
				$itemcount = 2
			}
			for ($i=0; $i -lt $itemcount; $i++){
				$newitem = $newitem.clone()
				$newitem.LineItemId = "$id"
				$id++
#### problem : wenn nur ein item funktioniert item[0] nicht !
				if(!$single_item_test){
				if($child.OrderReport.Item.SKU -eq "PCE-FWS 20 - Versand PCE Group"){
					$newitem.Article.BuyerAid = "PCE-FWS 20"
				}else{
					$newitem.Article.BuyerAid = $child.OrderReport.Item.SKU
				}
				$bestellbestaetigung.AmazonEnvelope.Message.OrderAcknowledgement.Item.AmazonOrderItemCode = $child.OrderReport.Item.AmazonOrderItemCode
				$content += "Artikel: "
				$content += $child.OrderReport.Item.SKU
				$content += "     Anzahl: "
				$content += $child.OrderReport.Item.Quantity
				$content += " `n"
				$newitem.Quantity = $child.OrderReport.Item.Quantity
				[float]$price = ([math]::round(($child.OrderReport.Item.ItemPrice.Component[0].Amount.Innertext/119*100),2) -replace "," ,".")
				$pricetemp = ($price / $child.OrderReport.Item.Quantity)
				$newitem.Price.PriceAmount = ($pricetemp.tostring() -replace "," ,".")
				[float]$priceamount = (($price -replace "," ,".")/[int]$child.OrderReport.Item.Quantity) ###### 28.07
				[float]$total += $price
				$itemnum = "1"
			}else{
				if($child.OrderReport.Item[$i].SKU -eq "PCE-FWS 20 - Versand PCE Group"){
					$newitem.Article.BuyerAid = "PCE-FWS 20"
				}else{
					$newitem.Article.BuyerAid = $child.OrderReport.Item[$i].SKU
				}
				$bestellbestaetigung.AmazonEnvelope.Message.OrderAcknowledgement.Item[$i].AmazonOrderItemCode = $child.OrderReport.Item.AmazonOrderItemCode
				$newitem.Quantity = $child.OrderReport.Item[$i].Quantity
				[float]$price = $child.OrderReport.Item[$i].ItemPrice.Component[0].Amount.Innertext/119*100
				$newitem.Price.PriceAmount = $price
				[float]$priceamount = ($price / $child.OrderReport.Item[$i].Quantity)   ### 28.07.10
				[float]$total += $price # = zu +=   // 19.07.10
				$itemnum = $itemnum + $child.OrderReport.Item[$i].Quantity
			}
				$newitem.Unit.Symbol = "Stck."
				$newitem.Unit.Code = "c62"
				$newitem.Price.PriceAmountTotal = $price.tostring()    # $total.tostring() -replace "," ,"."     // 19.07.10 - preise wurden "doppelt verdoppelt"
				$newitem.Price.PriceLineAmount = $price.tostring()     #$total.tostring() -replace "," ,"."		 // 19.07.10 - preise wurden "doppelt verdoppelt"
				$xml_order.order.itemlist.AppendChild($newitem) > $null
			}
			$xml_order.order.itemlist.item | Where-Object { $_.LineItemId -eq "" } |ForEach-Object  { [void]$xml_order.order.itemlist.RemoveChild($_) }
			$xml_order.order.OrderSummary.TotalItemNum = "$itemnum"
			$xml_order.order.OrderSummary.TotalAmount = $total.tostring() -replace "," ,"."
			$content += "Gesamtpreis ohne Versand: $priceamount `n"
			$content += $istkundeerstellt
			$content += " `n"
			$xml_order.Save("D:\import\amazon\ausgang\order\temp\order-$file_part.xml")
			$xml_order.Save("D:\import\amazon\ausgang\backup\order\order-$file_part.xml")
			$bestellbestaetigung.Save("D:\import\amazon\amazon-amtu\DocumentTransport\production\outgoing\order-acknowlegement-$file_part.xml")
			echo "saving file order-$file_part.xml"
		}
    move-item D:\import\amazon\amazon-amtu\DocumentTransport\production\reports\$datei D:\import\amazon\amazon-amtu\DocumentTransport\production\archiv\reports\
	} 
}
	#Throw
    echo "sleeping..."
    sleep 400
	#eparcel start
	cd "D:\import\amazon\ps\"
	php eParcel_check.php
	sleep 20


$dateinamen = get-ChildItem D:\import\amazon\ausgang\order\temp\*.* -name -include *.xml

foreach ($datei in $dateinamen)
	{
    if ($datei)
        {
        move-item D:\import\amazon\ausgang\order\temp\$datei D:\import\amazon\ausgang\order\

        }
    }
        
if ($process -eq 1)
{
sendmail "gsc@warensortiment.de" $content
sendmail "pph@warensortiment.de" $content
sendmail "fes@warensortiment.de" $content
}