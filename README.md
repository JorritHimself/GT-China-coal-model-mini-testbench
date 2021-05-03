# GT-China-coal-model-mini-testbench
A testbench version of a China coal transport model with massively simplified input data

# What it does
The code creates a linear optimization problem file, using the xlsx files in the folder /input  
The model optimizes for a minimum cost of production + transport + transmission.  
Production meaning coal mining costs, transport meaning rail/truck/riverborne/ocean-going transport and handling costs, and transmission meaning inter-provincial transport of electricity via UHV cables.

# Constraints
The constraints in the mini testbench are the same as in the full model.  
They are:
- Mines (or any other node) cannot supply types of coal they do not produce.
- The flow of coal of each type out of a node cannot exceed flows of coal of each type into a node plus supply by the node (with supply being non-zero only for mines).
- The energy content of the supply and the flows of coal of each type into a node have to be at least equal to the demand for electricity, plus other thermal coal demand, plus the energy content of flows of coal of each type out of a node. Note that only mines can supply coal, all demand for electrical power occurs in provincial demand nodes, and demand for other thermal coal is placed at city-level nodes.
- The amount of hard coking coal (HCC) flowing into a node has to at least be equal to the steel demand multiplied by 0.599. Note that all steel demand is placed in provincial level steel demand nodes, which are connecte with uni-directional links from steel plants to steel demand nodes. This means no coal can flow out of a steel demand node and we do not need further forumalae for mass balances. Also note that we presume a mix of coking coal need to produce a ton of steel of 599 kg Hard coking coal (HCC), 182 kg of soft coking coal (SCC), and 185 kg of pulverized coal for injection (PCI).
- The amount of soft coking coal (SCC) flowing into a node has to at least be equal to the amount of HCC flowing into that node, mulitplied with 0.599/0.182.
- The amount of pulverized coal for injection (PCI) flowing into a node has to at least be equal to the amount of HCC flowing into that node, mulitplied with 0.599/0.185.
- The total volume of all coal types transported along a link cannot exceed the transport capacity of that link. Note that this constraint is applied only to those links with a non-infinite transport capacity. In practive this means rail links are assumed to have a transport capacity, ocean routes, rivers, and road links are assumed to have infinite capacity.
- The total amount of energy transported along a link cannot exceed the transmission capacity of that link. That is, the amount of each coal type multiplied with the energy content of each coal type cannot exceed the electrical transmission capacity of links. This constraint is applied only links between power plant units and provincial electricity demand nodes, as well as UHV transmission links between provincial electricity demand nodes. These are the only links along which electrical energy is transported. All other links transport physcial quantities of coal. This line simultaneously deals with the production capapcity (MW) and conversion efficiency of power plants: the energy transported over a link cannot exceed the volume of each coal type, multiplied with the energy content of each coal type, multiplied with the energy conversion factor of the link. For links between coal fired power plant units and provincial electricity demand nodes, this is equal to the conversion effincy of the power plant unit. For UHV transmission links between two provincial level electricity demand nodes, this is equal to (1- transmission losses) over that UHV line, with transmission losses calculated based on transmission distance and a benchmark loss for UHV-DC or UHV-AC lines.
- The handling capacity of ports cannot be exceeded. Specifically, the total amount of coal flowing out of a port cannot exceed its handling capacity.
- The production capacity of steel plants cannot be exceeded. Specifically, the total amount of hard coking coal, soft coking coal, and pulverized coal for injection flowing out of a steel plant node (and towards a provincial steel demand node) cannot exceed the steel plant's production capacity multiplied by 0.599+0.182+0.185, the mix of different coking coals needed to produce steel.

# Notes
- All transport costs are pre-calculated for each link, and include a fixed handling costs and a distance based transport cost, based on the type handling (origin and destaination) and type of transport (separate for rail, truck, riverborne, ocean-going. A small number of coal rail lines has specific handling and transport costs).
- Some of the capacities are already reported in the input sheet for the edges. The physical transport capacity from this sheet is used. For capacities of ports, steel plants, and electrical transmission capacities, the data from the separate port/steel plant/electrical capacities sheets is used.

# Credit
This project was developed by:
- Jorrit Gosens: conceptualization, bulk of the data preparation, and software implementation and debugging in R & Python;
- Alex H. Turnbull: conceptualization, some data preparation, and proof-reading and debugging of software.

# License
MIT License as separately included.  
In short, do what you want with this script, but refer to the original authors when you use or develop this code.
