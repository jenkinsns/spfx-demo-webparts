
export class employee{
  private empcode:number;
  private empName:string;

  constructor(name:string, code:number)
  {
    this.empName = name;
    this.empcode = code;
    this.printval(this.empName,this.empcode);
  }
  private printval(name:string,code:number)
  {
    console.log("Employee Name : " + name + " Employee Code : "+ code);
  }

  public printfromfunction(name:string,code:number)
  {
    return "Welcome : " + name + " your code :" + code;
  }
}
