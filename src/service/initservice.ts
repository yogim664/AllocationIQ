import { WebPartContext } from "@microsoft/sp-webpart-base";
import { graphfi, GraphFI, SPFx as graphSPFx } from "@pnp/graph";
import { ISPFXContext, spfi, SPFI, SPFx as spSPFx } from "@pnp/sp";
 
let _sp: SPFI;
let _graph: GraphFI;
 
export const initService = (context: WebPartContext): void => {
  _sp = spfi().using(spSPFx(context as ISPFXContext));
  _graph = graphfi().using(graphSPFx(context as ISPFXContext));
};
 
export const getSP = (): SPFI => {
  if (!_sp) throw new Error("SP not initialized");
  return _sp;
};
 
export const getGraph = (): GraphFI => {
  if (!_graph) throw new Error("Graph not initialized");
  return _graph;
};
 
 