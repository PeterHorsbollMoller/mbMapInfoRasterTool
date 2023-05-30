using System;
using MapInfo.Events;
using MapInfo.Types;

namespace MapInfoProAdvancedTool
{
    /*public class MapInfoProAddInBase : IMapInfoProAddInInterface
    {
        protected IMapInfoPro _application;
        protected IMapInfoEvents _events;

        /// <summary>
        /// 
        /// </summary>
        public IMapBasicApplication CurrentInstance
        {
            get;
            set;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="app"></param>
        /// <param name="addInName"></param>
        public virtual void Initialize(IMapInfoPro app, string addInName)
        {
            _application = app;
            CurrentInstance = _application.GetMapBasicApplication(addInName);
            _events = MapInfoEvents.InitializeEvents(app, CurrentInstance);

            if (!UriParser.IsKnownScheme("pack"))
            {
                UriParser.Register(new GenericUriParser(GenericUriParserOptions.GenericAuthority), "pack", -1);
            }


            _application.RunMapBasicCommand(string.Format("print \"App Domain Id: {0}\"", AppDomain.CurrentDomain.Id), false);
            _application.RunMapBasicCommand(string.Format("print \"AddIn Name: {0}\"", CurrentInstance.Name), false);
        }

        /// <summary>
        /// 
        /// </summary>
        public virtual void Unload()
        {
        }
    }*/
}
