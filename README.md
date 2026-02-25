# TubeTabSlot — SolidWorks Add-in

SolidWorks add-in that creates tab & slot features on intersecting weldment tubes. Select two tubular weldment bodies, click the toolbar button, and the add-in generates extruded tabs and matching slot cuts at the intersection.

## Usage

1. Open a part with weldment tube bodies
2. Select two bodies — first is the **tab** tube, second is the **slot** tube
3. Click **Tab & Slot** in the toolbar (or menu)
4. Choose placement (both sides, near only, or far only) and tab depth
5. Click OK

## Build

Requires Visual Studio 2017 Build Tools (or later) and SolidWorks installed.

```
MSBuild.exe TubeTabSlot.csproj /p:Configuration=Debug /nologo
```

Build from an **admin** command prompt — the post-build step runs `regasm /codebase` to register the COM add-in.

## Registration

After building, restart SolidWorks. The add-in appears in **Tools → Add-Ins** as "Tab & Slot". Enable it and the toolbar button appears in Part documents.

Re-registration is only needed if you change the GUID, move the DLL, or switch between Debug/Release.

## Project Structure

| File | Purpose |
|------|---------|
| `SwAddin.cs` | Add-in entry point, command manager, toolbar setup |
| `TabSlotPMP.cs` | Property Manager Page handler (UI options) |
| `TabSlotRunner.cs` | Orchestrates feature creation |
| `SelectionHelper.cs` | Validates and reads the two-body selection |
| `GeometryHelper.cs` | Intersection geometry, plane computation, vector math |
| `TabSketchHelper.cs` | Draws tab cross-section profiles on sketch planes |
| `FeatureHelper.cs` | Extrude tab boss and cut slot operations |
| `DebugHelper.cs` | Debug visualization (behind `#if DEBUG`) |
