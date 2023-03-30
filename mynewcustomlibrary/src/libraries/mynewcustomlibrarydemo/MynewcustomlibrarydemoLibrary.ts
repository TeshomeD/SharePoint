export class MynewcustomlibrarydemoLibrary {
  public name(): string {
    return 'MynewcustomlibrarydemoLibrary';
  }
  
  public getCurrentTime(): string{
    let currentDate: Date;
    let str:string;

    currentDate = new Date();
    str= "<br> Todays date is: "+ currentDate.toDateString();
    str+= "<br> Current time is: "+ currentDate.toTimeString();

    return (str);
  }
}
