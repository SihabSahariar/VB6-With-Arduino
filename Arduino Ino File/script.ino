//VB6 GUI APP FOR ARDUINO - (13 FEB 2021)
//Sihab Sahariar
int LED = 13; // You may change
String text;

void setup()
{

  Serial.begin(9600);
  pinMode(LED, OUTPUT);
}

void loop()
{
	 if(Serial.avilable()>0)
	 {
		 text = Serial.readString();
		 if(text=="LEDON")
		 {
			 digitalWrite(LED,HIGH);
			 Serial.println("LED IS ON");
		 }
		 else if(text=="LEDOFF")
		 {
			 digitalWrite(LED,LOW);
			 Serial.println("LED IS OFF");
		 }
	}
	 
  
}

