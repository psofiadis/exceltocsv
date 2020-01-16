package eu.rea.ec.bo;

import eu.rea.ec.bo.worker.ToCSVTransformer;

public class ExcelToCSVUtil {

  public static void main(String[] args){
    new ToCSVTransformer().execute();
  }
}
