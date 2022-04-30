using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateRoomsPlugin
{
    [Transaction(TransactionMode.Manual)]
    public class CreateRooms : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            List<Level> levels = new FilteredElementCollector(doc)
                                    .OfClass(typeof(Level))
                                    .OfType<Level>()
                                    .ToList();

            View activeView = doc.ActiveView;
            Parameter parameter = activeView.get_Parameter(BuiltInParameter.VIEW_PHASE);
            ElementId eID = parameter.AsElementId();
            Phase phase = doc.GetElement(eID) as Phase;

            Transaction transaction = new Transaction(doc, "Размещение помещений");
            transaction.Start();
            List<Room> rooms = InsertNewRoomInPlanCircuit(doc, levels, phase);
            transaction.Commit();

            return Result.Succeeded;
        }

        List<Room> InsertNewRoomInPlanCircuit(Document document, List<Level> levels, Phase newConstructionPhase)
        {
            List<Room> rooms = new List<Room>();
            Floor overlap = new FilteredElementCollector(document)
                                    .OfClass(typeof(Floor))
                                    .OfType<Floor>()
                                    .FirstOrDefault();
            double overlapWidth = UnitUtils.ConvertFromInternalUnits(overlap.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM).AsDouble(), UnitTypeId.Millimeters);

            for (int i = 0; i < levels.Count(); i++)
            {
                PlanTopology planTopology = document.get_PlanTopology(levels[i]);
                int roomNumber = 1;
                foreach (PlanCircuit circuit in planTopology.Circuits)
                {
                    Room newScheduleRoom = document.Create.NewRoom(newConstructionPhase);
                    newScheduleRoom.Name = null;
                    newScheduleRoom.Number = (i + 1).ToString() + "_" + roomNumber.ToString();
                    Room newRoom = document.Create.NewRoom(newScheduleRoom, circuit);
                    SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
                    newRoom.GetBoundarySegments(options).FirstOrDefault();
                    try
                    {
                        newRoom.get_Parameter(BuiltInParameter.ROOM_UPPER_LEVEL).Set(levels[i + 1].Id);
                        newRoom.get_Parameter(BuiltInParameter.ROOM_UPPER_OFFSET).Set(-UnitUtils.ConvertToInternalUnits(overlapWidth, UnitTypeId.Millimeters));
                    }
                    catch
                    {
                    }
                    rooms.Add(newRoom);
                    roomNumber++;
                }
            }
            return rooms;
        }
    }
}
