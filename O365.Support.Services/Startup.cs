using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.HttpsPolicy;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace O365.Support.Services
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddSwaggerGen(options =>
            {
                options.SwaggerDoc("v1", new Microsoft.OpenApi.Models.OpenApiInfo
                {
                    Title = "O365 Support Services API",
                    Version = "1.0",
                    Contact = new Microsoft.OpenApi.Models.OpenApiContact
                    {
                        Name = "Purusothaman Murugan",
                        Email = "purusothamanm@gmail.com"

                    },

                    Description = "API Created for Service Operations"
                });

            });
            services.AddControllers();
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            { 
                app.UseDeveloperExceptionPage();
            }

            app.UseHttpsRedirection(); 

            app.UseRouting();

            app.UseAuthorization();            

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllers();
            });

            app.UseSwagger();

            if (env.IsDevelopment())
            {

                app.UseSwaggerUI(options =>
                {
                    options.SwaggerEndpoint("/swagger/v1/swagger.json", "O365 Support Services API");
                    options.RoutePrefix = "";
                });

                //// Allowing requests from the client - APPLICATION LEVEL
                ////app.UseCors(builder =>
                ////    builder.WithOrigins("http://localhost:55294"));
                //app.UseCors(builder =>
                //                            builder
                //                            .AllowAnyOrigin()
                //                            .AllowAnyMethod()
                //                            .AllowAnyHeader());
                app.UseDeveloperExceptionPage();
            }
        }
    }
}
