import com.thingworx.logging.LogUtilities;
import com.thingworx.metadata.annotations.ThingworxServiceDefinition;
import com.thingworx.metadata.annotations.ThingworxServiceParameter;
import com.thingworx.metadata.annotations.ThingworxServiceResult;
import com.thingworx.resources.Resource;
import com.thingworx.types.InfoTable;
import org.slf4j.Logger;

public class ExcelReportGenerator extends Resource {

	private static Logger _logger = LogUtilities.getInstance().getApplicationLogger(ExcelReportGenerator.class);

	public ExcelReportGenerator() {
		// TODO Auto-generated constructor stub
	}

	@ThingworxServiceDefinition(name = "ExcelExport", description = "", category = "", isAllowOverride = false, aspects = {
			"isAsync:false" })
	@ThingworxServiceResult(name = "result", description = "Thingworx Download String\r\n", baseType = "STRING", aspects = {})
	public String ExcelExport(
			@ThingworxServiceParameter(name = "InfoTableData", description = "", baseType = "INFOTABLE", aspects = {
					"isRequired:true", "isEntityDataShape:true" }) InfoTable InfoTableData,
			@ThingworxServiceParameter(name = "RepositoryName", description = "", baseType = "THINGNAME", aspects = {
					"isRequired:true", "thingTemplate:FileRepository" }) String RepositoryName,
			@ThingworxServiceParameter(name = "ExcelTemplateLocation", description = "", baseType = "STRING") String ExcelTemplateLocation) {
		_logger.trace("Entering Service: ExcelExport");
		_logger.trace("Exiting Service: ExcelExport");
		return null;
	}

}
