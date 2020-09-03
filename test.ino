void setup() 
{
  Serial.begin(9600);
  Serial.print("date,time\n");
}

void loop() 
{
  Serial.print("225,144\n");
  delay(2000);
}
