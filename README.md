# Networking-Automation
Fetching data from Huawei routers and uploading in Excel sheet using Pyexcel library

""" This code will fetch data from Huawei Routers and store in excel file systmatically.
For IP's create file endpointIPs_1.txt and endpointIPs_2.txt and so on and call in function so you can test ping result for different destination addresses.
You can use this code for any vendor device.
Output will look like below in excel formate with different sheets for different input files.
Management IP	Hostname	                    GBN ID	    Model	Ping Status     Success Rate(%)	    LAN Subnet	        LAN Status
172.31.122.7	GBN-0416_PHO-SEC-SCH_AR12-01	GBN-0416	AR12	Pingable	    100	                10.191.32.1/30	        UP
172.31.108.9	                    Device is not reachable						
172.31.86.58	GBN-0620_RES-IND-PAR_AR12-01	GBN-0620	AR12	Not Pingable    Unsuccessfull	    10.70.36.1/22	        Down
172.31.53.216	GBN-1112_KWA-SEC-SCH_AR12-01	GBN-1112	AR12	Pingable        100	                10.206.56.1/21	        UP
172.31.82.37	GBN-0456_LUK-PRI-SCH_AR22-01	GBN-0456	AR22	Not Pingable	Unsuccessfull	    10.191.184.1/21	        Down
172.31.108.26	                    Device is not reachable						
172.31.78.29	GBN-0444_GST-GRN-DAV_AR22-01	GBN-0444	AR22	Not Pingable	Unsuccessfull	    10.2.2.1/24	            UP
172.31.78.31	GBN-0446_GST-02F-GIF_AR12-01	GBN-0446	AR12	Pingable	    100	                10.10.0.3/24	        UP
172.31.108.16	                    Device is not reachable						
172.31.90.211	GBN-0575_EAS-DRI-LIB_AR12-01	GBN-0575	AR12	Pingable	    100	                10.6.14.1/24	        UP
Contact me if need any assistance on advit.raut@gmail.com
    """

